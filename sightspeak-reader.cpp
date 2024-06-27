#include <iostream>
#include <UIAutomation.h>
#include <windows.h>
#include <tchar.h>
#include <chrono>
#include <mutex>
#include <queue>
#include <thread>
#include <condition_variable>
#include <atlbase.h>
#include <crtdbg.h>
#include <fcntl.h>
#include <io.h>
#include <algorithm>
#include <sapi.h>
#include <atomic>
#include <unordered_set>
#include <sstream>
#include <fstream>
#include <future>
#include <shared_mutex>

// Utility functions for string conversion
std::wstring Utf8ToWstring(const std::string& str) {
    if (str.empty()) return std::wstring();
    int size_needed = MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), NULL, 0);
    std::wstring wstrTo(size_needed, 0);
    MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), &wstrTo[0], size_needed);
    return wstrTo;
}

std::string WstringToUtf8(const std::wstring& wstr) {
    if (wstr.empty()) return std::string();
    int size_needed = WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), NULL, 0, NULL, NULL);
    std::string strTo(size_needed, 0);
    WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), &strTo[0], size_needed, NULL, NULL);
    return strTo;
}

// Structure to hold text and its corresponding rectangle
struct TextRect {
    std::wstring text;
    RECT rect;
};

// Mouse hook handle
HHOOK hMouseHook;

// UI Automation interface pointers
CComPtr<IUIAutomation> pAutomation = NULL;
CComPtr<IUIAutomationElement> pPrevElement = NULL;
RECT prevRect = { 0, 0, 0, 0 };

// Time point for last processed event
std::chrono::time_point<std::chrono::steady_clock> lastProcessedTime;

// Mutex for general protection
std::mutex mtx;

// Queue for storing points (cursor positions)
std::queue<POINT> pointsQueue;
std::mutex queueMtx; // Mutex for queue protection
std::condition_variable cv; // Condition variable for queue operations
std::atomic<bool> stopWorker(false); // Flag to stop worker threads

// Speech synthesis interface pointer
std::mutex pVoiceMutex;
CComPtr<ISpVoice> pVoice = NULL;

// Atomic flags and time points for cursor movement
std::atomic<bool> cursorMoving(false);
std::chrono::time_point<std::chrono::steady_clock> lastCursorMoveTime;

// Atomic flag to stop processing
std::atomic<bool> stopProcessing(false);

// Mutex and condition variable for speech operations
std::mutex speechMtx;
std::condition_variable speechCv;
std::queue<TextRect> speechQueue;
std::atomic<bool> newElementDetected(false); // Flag to indicate new element detection

// Thread and flags for speech processing
std::thread speechThread;
std::atomic<bool> speaking(false);
std::condition_variable speakingCv;

// Atomic flag to indicate if a rectangle is drawn
std::atomic<bool> rectangleDrawn(false);

// Constants for processing limits
const size_t MAX_TEXT_LENGTH = 10000;
const int MAX_DEPTH = 5;
const int MAX_CHILDREN = 20;
const int MAX_ELEMENTS = 200;

// Time point for last rectangle drawing event
std::chrono::time_point<std::chrono::steady_clock> lastRectangleDrawTime;

// Shared state for rectangle management
std::shared_mutex rectQueueMtx;
std::atomic<bool> speechOngoing(false);
RECT currentRect;
std::mutex rectMtx;
std::condition_variable rectCv;

// Enum for rectangle actions
enum class RectangleAction { Draw, Clear };

// Structure to hold rectangle tasks
struct RectangleTask {
    RECT rect;
    RectangleAction action;
};

// Queue for rectangle tasks
std::queue<RectangleTask> rectQueue;
std::mutex rectQueueMtx; // Mutex for rectangle queue protection
std::condition_variable rectQueueCv; // Condition variable for rectangle queue operations
std::atomic<bool> stopRectWorker(false); // Flag to stop rectangle worker threads

// Global mutex for thread-safe logging
std::mutex logMutex;

// Template function to submit a task to be run asynchronously
template <typename F, typename... Args>
auto submitTask(F&& f, Args&&... args) {
    // Use std::async to run the task asynchronously with std::launch::async policy
    // std::launch::async ensures the task is run in a new thread
    // std::forward is used to perfectly forward the function and arguments to preserve their value categories (lvalue/rvalue)
    return std::async(std::launch::async, std::forward<F>(f), std::forward<Args>(args)...);
}

void DebugLog(const std::wstring& message) {
    // Create a wide string stream for the debug message
    std::wstringstream ws;
    ws << L"[DEBUG] " << message << std::endl;
    // Output the debug message to the debugger
    OutputDebugString(ws.str().c_str());
    // Lock the mutex for thread-safe file access
    std::lock_guard<std::mutex> guard(logMutex);
    // Open the log file in append mode
    std::wofstream logFile("debug.log", std::ios::app);
    // Check if the file is open
    if (logFile.is_open()) {
        // Write the debug message to the log file
        logFile << ws.str();
        // Ensure the message is flushed to the file
        logFile.flush();
    }
    else {
        // Handle the error (optional)
        OutputDebugString(L"Failed to open debug.log\n");
    }
}

void SetConsoleBufferSize() {
    // Initialize CONSOLE_SCREEN_BUFFER_INFOEX structure
    CONSOLE_SCREEN_BUFFER_INFOEX csbiex = { 0 };
    csbiex.cbSize = sizeof(csbiex);
    // Get the handle to the console output buffer
    HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
    // Retrieve current console screen buffer information
    if (GetConsoleScreenBufferInfoEx(hConsole, &csbiex)) {
        // Set the buffer size dimensions
        csbiex.dwSize.X = 8000;
        csbiex.dwSize.Y = 2000;
        // Apply the new console screen buffer size
        SetConsoleScreenBufferInfoEx(hConsole, &csbiex);
    }
}

void QueueRectangle(const RECT& rect, RectangleAction action) {
    {
        // Unique lock for writing to the queue
        std::unique_lock<std::shared_mutex> lock(rectQueueMtx);

        // Add the rectangle task to the queue
        rectQueue.push({ rect, action });

        // Log the action and rectangle coordinates
        std::wstring actionStr = (action == RectangleAction::Draw ? L"Queueing rectangle for drawing: (" : L"Queueing rectangle for clearing: (");
        DebugLog(actionStr + std::to_wstring(rect.left) + L", " +
            std::to_wstring(rect.top) + L", " +
            std::to_wstring(rect.right) + L", " +
            std::to_wstring(rect.bottom) + L")");
    }

    // Notify one waiting thread that a new task is available
    rectQueueCv.notify_one();
}


void ProcessRectangle(const RECT& rect, RectangleAction action) {
    // Get the device context for the entire screen
    HDC hdc = GetDC(NULL);
    if (hdc) {
        if (action == RectangleAction::Draw) {
            // Create a red pen for drawing the rectangle
            HPEN hPen = CreatePen(PS_SOLID, 2, RGB(255, 0, 0));
            HGDIOBJ hOldPen = SelectObject(hdc, hPen);
            // Select a hollow brush (no fill) for the rectangle
            HBRUSH hBrush = (HBRUSH)GetStockObject(HOLLOW_BRUSH);
            HGDIOBJ hOldBrush = SelectObject(hdc, hBrush);
            // Draw the rectangle
            Rectangle(hdc, rect.left, rect.top, rect.right, rect.bottom);
            // Restore the original pen and brush
            SelectObject(hdc, hOldPen);
            SelectObject(hdc, hOldBrush);
            // Delete the created pen
            DeleteObject(hPen);
        }
        else {
            // Invalidate the rectangle area to force a redraw
            InvalidateRect(NULL, &rect, TRUE);
            // Log the invalidation action
            DebugLog(L"Invalidated rectangle area: (" + std::to_wstring(rect.left) + L", " +
                std::to_wstring(rect.top) + L", " + std::to_wstring(rect.right) + L", " +
                std::to_wstring(rect.bottom) + L")");
        }
        // Release the device context
        ReleaseDC(NULL, hdc);
    }
    else {
        // Log if getting the device context fails
        DebugLog(L"Failed to get device context.");
    }
}

void ProcessRectangleAsync(const RECT& rect, RectangleAction action) {
    // Submit the ProcessRectangle function to be run asynchronously with the provided rect and action
    submitTask(ProcessRectangle, rect, action);
}

void RectangleWorker() {
    DebugLog(L"Starting RectangleWorker.");

    while (!stopRectWorker.load(std::memory_order_acquire)) {
        // Unique lock for accessing the queue
        std::unique_lock<std::shared_mutex> lock(rectQueueMtx);

        // Wait for the condition variable notification or stop signal
        rectQueueCv.wait(lock, [] { return !rectQueue.empty() || stopRectWorker.load(std::memory_order_acquire); });

        if (stopRectWorker.load(std::memory_order_acquire)) {
            DebugLog(L"Stopping RectangleWorker.");
            break;
        }
        // Process all tasks in the rectangle queue
        while (!rectQueue.empty()) {
            // Use unique lock to modify the queue
            std::unique_lock<std::shared_mutex> uniqueLock(rectQueueMtx);
            // Get the task from the front of the queue and remove it
            RectangleTask task = rectQueue.front();
            rectQueue.pop();
            // Unlock immediately after modifying the queue
            uniqueLock.unlock();
            // Log the rectangle being processed
            DebugLog(L"Processing rectangle: (" + std::to_wstring(task.rect.left) + L", " +
                std::to_wstring(task.rect.top) + L", " +
                std::to_wstring(task.rect.right) + L", " +
                std::to_wstring(task.rect.bottom) + L")");
            // Process the rectangle asynchronously
            ProcessRectangleAsync(task.rect, task.action);
        }
    }

    DebugLog(L"Exiting RectangleWorker.");
}

void PrintText(const std::wstring& text) {
    //Prints all text from ProrocessCursorPosition
    std::wcout << text << std::endl;
    std::wcout.flush();
    DebugLog(L"Printed text to console: " + text);
}

void StopSpeechAndClearQueue() {
    DebugLog(L"Entering StopSpeechAndClearQueue.");
    std::unique_lock<std::mutex> speechLock(speechMtx);
    // Stop ongoing speech
    {
        std::lock_guard<std::mutex> lock(pVoiceMutex); // Lock for thread-safe access to pVoice
        if (pVoice) {
            HRESULT hr = pVoice->Speak(nullptr, SPF_PURGEBEFORESPEAK, nullptr);
            if (FAILED(hr)) {
                DebugLog(L"Failed to stop speech: " + std::to_wstring(hr));
            }
            else {
                DebugLog(L"Successfully stopped speech.");
            }
        }
    }
    // Clear the speech queue
    DebugLog(L"Speech queue size before clearing: " + std::to_wstring(speechQueue.size()));
    std::queue<TextRect> emptyQueue;
    std::swap(speechQueue, emptyQueue);
    DebugLog(L"Cleared speech queue.");
    // Wait until the speaking flag is false
    speakingCv.wait(speechLock, [] { return !speaking.load(std::memory_order_acquire); });
    DebugLog(L"speaking flag is now false.");
    DebugLog(L"Exiting StopSpeechAndClearQueue.");
}

void StopCurrentProcesses() {
    DebugLog(L"Entering StopCurrentProcesses.");
    try {
        // Stop ongoing speech and clear the speech queue
        StopSpeechAndClearQueue();

        RECT tempRect;
        {
            // Lock mutex to safely access and copy currentRect
            std::lock_guard<std::mutex> rectLock(rectMtx);
            tempRect = currentRect;
            // Log the current rectangle to be cleared
            DebugLog(L"Current rectangle to be cleared: (" + std::to_wstring(tempRect.left) + L", " +
                std::to_wstring(tempRect.top) + L", " +
                std::to_wstring(tempRect.right) + L", " +
                std::to_wstring(tempRect.bottom) + L")");
        }
        // Queue the rectangle for clearing
        QueueRectangle(tempRect, RectangleAction::Clear);
    }
    catch (const std::system_error& e) {
        // Log system errors
        DebugLog(L"Exception in StopCurrentProcesses: " + Utf8ToWstring(e.what()));
    }
    catch (const std::exception& e) {
        // Log other exceptions
        DebugLog(L"Exception in StopCurrentProcesses: " + Utf8ToWstring(e.what()));
    }
    DebugLog(L"Exiting StopCurrentProcesses.");
}

void DebouncedStopCurrentProcesses() {
    static std::chrono::steady_clock::time_point lastCallTime = std::chrono::steady_clock::now();
    static std::mutex debounceMutex; // Mutex for thread safety
    auto now = std::chrono::steady_clock::now();
    std::lock_guard<std::mutex> lock(debounceMutex); // Lock for thread-safe access
    // Check if more than 300 milliseconds have passed since the last call
    if (std::chrono::duration_cast<std::chrono::milliseconds>(now - lastCallTime).count() > 300) {
        lastCallTime = now; // Update the last call time
        StopCurrentProcesses(); // Call the function to stop current processes
    }
}

void SpeakTextAsync(const std::wstring& textToSpeak) {
    submitTask([=]() {
        std::lock_guard<std::mutex> lock(pVoiceMutex); // Lock for thread-safe access to pVoice
        if (pVoice) {
            HRESULT hr = pVoice->Speak(nullptr, SPF_PURGEBEFORESPEAK, nullptr);
            if (SUCCEEDED(hr)) {
                hr = pVoice->Speak(textToSpeak.c_str(), SPF_ASYNC, nullptr);
                if (SUCCEEDED(hr)) {
                    DebugLog(L"Successfully spoke text: " + textToSpeak);
                }
                else {
                    DebugLog(L"Failed to speak text: " + textToSpeak + L" Error: " + std::to_wstring(hr));
                }
            }
            else {
                DebugLog(L"Failed to purge speech: " + std::to_wstring(hr));
            }
        }
        });
}

void SpeakText() {
    DebugLog(L"SpeakText thread started.");
    try {
        while (!stopWorker.load(std::memory_order_acquire)) {
            std::unique_lock<std::mutex> lock(speechMtx);
            DebugLog(L"Waiting on speechCv in SpeakText.");
            speechCv.wait(lock, [] { return !speechQueue.empty() || stopWorker.load(std::memory_order_acquire); });
            DebugLog(L"Finished waiting on speechCv in SpeakText.");

            if (stopWorker.load(std::memory_order_acquire)) {
                DebugLog(L"Stopping speech thread.");
                return;
            }

            while (!speechQueue.empty()) {
                TextRect textRectPair = speechQueue.front();
                speechQueue.pop();

                const std::wstring& textToSpeak = textRectPair.text;
                const RECT& textRect = textRectPair.rect;

                if (textToSpeak.empty()) {
                    DebugLog(L"Text to speak is empty, skipping.");
                    continue;
                }

                DebugLog(L"Attempting to speak text: " + textToSpeak);
                speaking.store(true, std::memory_order_release);

                SpeakTextAsync(textToSpeak);

                {
                    std::lock_guard<std::mutex> rectLock(rectMtx);
                    currentRect = textRect;
                }
                QueueRectangle(currentRect, RectangleAction::Draw);

                SPVOICESTATUS status;
                while (speaking.load(std::memory_order_acquire)) {
                    if (pVoice) {
                        HRESULT hr = pVoice->GetStatus(&status, nullptr);
                        if (FAILED(hr)) {
                            DebugLog(L"Failed to get status: " + std::to_wstring(hr));
                            break;
                        }

                        if (status.dwRunningState == SPRS_DONE) {
                            DebugLog(L"Speech done for text: " + textToSpeak);
                            break;
                        }
                    }

                    std::this_thread::sleep_for(std::chrono::milliseconds(100));
                }

                QueueRectangle(currentRect, RectangleAction::Clear);

                speaking.store(false, std::memory_order_release);
                speakingCv.notify_all();
                DebugLog(L"Finished processing text: " + textToSpeak);
            }
        }
    }
    catch (const std::system_error& e) {
        DebugLog(L"Exception in SpeakText: " + Utf8ToWstring(e.what()));
    }
    catch (const std::exception& e) {
        DebugLog(L"Exception in SpeakText: " + Utf8ToWstring(e.what()));
    }
}

void StartSpeechThread() {
    stopWorker.store(false, std::memory_order_release);
    speechThread = std::thread(SpeakText);
    DebugLog(L"Started speech thread.");
}

void StopSpeechThread() {
    {
        std::lock_guard<std::mutex> lock(speechMtx);
        stopWorker.store(true, std::memory_order_release);
    }
    speechCv.notify_all();
    if (speechThread.joinable()) {
        speechThread.join();
    }
    DebugLog(L"Stopped speech thread.");
}

bool IsDifferentElement(CComPtr<IUIAutomationElement> pElement) {
    DebugLog(L"Attempting to lock mtx in IsDifferentElement.");
    std::lock_guard<std::mutex> lock(mtx);
    DebugLog(L"Locked mtx in IsDifferentElement.");
    if (pPrevElement == NULL && pElement == NULL) {
        return false;
    }
    else if (pPrevElement == NULL || pElement == NULL) {
        return true;
    }

    BOOL areSame;
    HRESULT hr = pAutomation->CompareElements(pPrevElement, pElement, &areSame);
    DebugLog(L"Compared elements in IsDifferentElement.");
    return SUCCEEDED(hr) && !areSame;
}

void ReinitializeAutomation() {
    DebugLog(L"Reinitializing UI Automation and COM components");

    pAutomation.Release();
    pPrevElement.Release();
    pVoice.Release();

    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) {
        DebugLog(L"Failed to reinitialize COM library: " + std::to_wstring(hr));
        return;
    }

    hr = CoCreateInstance(__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pAutomation));
    if (FAILED(hr)) {
        DebugLog(L"Failed to reinitialize UI Automation: " + std::to_wstring(hr));
        CoUninitialize();
        return;
    }

    hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void**)&pVoice);
    if (FAILED(hr)) {
        DebugLog(L"Failed to reinitialize SAPI: " + std::to_wstring(hr));
        pAutomation.Release();
        CoUninitialize();
        return;
    }

    pVoice->SetVolume(100);
    pVoice->SetRate(0);

    DebugLog(L"Successfully reinitialized UI Automation and COM components");
}

void ReadElementText(CComPtr<IUIAutomationElement> pElement, std::queue<TextRect>& textRectQueue, size_t& currentTextLength, std::unordered_set<std::wstring>& processedTexts) {
    if (stopProcessing.load(std::memory_order_acquire)) return;

    try {
        CComPtr<IUIAutomationTextPattern> pTextPattern = NULL;
        HRESULT hr = pElement->GetCurrentPatternAs(UIA_TextPatternId, IID_PPV_ARGS(&pTextPattern));
        if (SUCCEEDED(hr) && pTextPattern) {
            CComPtr<IUIAutomationTextRange> pTextRange = NULL;
            hr = pTextPattern->get_DocumentRange(&pTextRange);
            if (SUCCEEDED(hr) && pTextRange) {
                CComBSTR text;
                hr = pTextRange->GetText(-1, &text);
                if (SUCCEEDED(hr)) {
                    std::wstring textStr(static_cast<wchar_t*>(text));
                    if (!textStr.empty() && processedTexts.find(textStr) == processedTexts.end()) {
                        processedTexts.insert(textStr);
                        currentTextLength += textStr.length();
                        if (currentTextLength <= MAX_TEXT_LENGTH) {
                            RECT rect = {};
                            SAFEARRAY* pRects = NULL;
                            hr = pTextRange->GetBoundingRectangles(&pRects);
                            if (SUCCEEDED(hr) && pRects) {
                                LONG lBound = 0, uBound = 0;
                                SafeArrayGetLBound(pRects, 1, &lBound);
                                SafeArrayGetUBound(pRects, 1, &uBound);
                                for (LONG i = lBound; i <= uBound; i += 4) {
                                    double rectArr[4];
                                    SafeArrayGetElement(pRects, &i, &rectArr);
                                    rect.left = static_cast<LONG>(rectArr[0]);
                                    rect.top = static_cast<LONG>(rectArr[1]);
                                    rect.right = static_cast<LONG>(rectArr[2]);
                                    rect.bottom = static_cast<LONG>(rectArr[3]);
                                }
                                SafeArrayDestroy(pRects);
                            }
                            textRectQueue.push({ textStr, rect });
                            DebugLog(L"Read text: " + textStr);
                        }
                        else {
                            DebugLog(L"Text length exceeded maximum allowed length.");
                        }
                    }
                    else {
                        DebugLog(L"Text already processed or empty.");
                    }
                }
                else {
                    DebugLog(L"Failed to get text from text range: " + std::to_wstring(hr));
                    if (hr == UIA_E_ELEMENTNOTAVAILABLE) {
                        DebugLog(L"Element not available. Skipping element.");
                    }
                }
            }
            else {
                DebugLog(L"Failed to get document range: " + std::to_wstring(hr));
            }
        }
        else {
            if (hr != UIA_E_NOTSUPPORTED) {
                CComBSTR name;
                hr = pElement->get_CurrentName(&name);
                if (SUCCEEDED(hr)) {
                    std::wstring nameStr(static_cast<wchar_t*>(name));
                    if (!nameStr.empty() && processedTexts.find(nameStr) == processedTexts.end()) {
                        processedTexts.insert(nameStr);
                        currentTextLength += nameStr.length();
                        if (currentTextLength <= MAX_TEXT_LENGTH) {
                            RECT rect = {};
                            hr = pElement->get_CurrentBoundingRectangle(&rect);
                            if (SUCCEEDED(hr)) {
                                textRectQueue.push({ nameStr, rect });
                                DebugLog(L"Read element name: " + nameStr);
                            }
                            else {
                                DebugLog(L"Failed to get bounding rectangle.");
                            }
                        }
                        else {
                            DebugLog(L"Element name length exceeded maximum allowed length.");
                        }
                    }
                    else {
                        DebugLog(L"Element name already processed or empty.");
                    }
                }
                else {
                    DebugLog(L"Failed to get element name: " + std::to_wstring(hr));
                }
            }
            else {
                DebugLog(L"Element does not support TextPattern: " + std::to_wstring(hr));
            }
        }
    }
    catch (const std::exception& e) {
        DebugLog(L"Exception in ReadElementText: " + Utf8ToWstring(e.what()));
    }
}

void CollectElementsDFS(CComPtr<IUIAutomationElement> pElement, std::vector<CComPtr<IUIAutomationElement>>& elements, int maxDepth, int maxChildren, int maxElements) {
    if (!pElement) return;

    struct ElementInfo {
        CComPtr<IUIAutomationElement> element;
        int depth;
    };

    std::queue<ElementInfo> elementQueue;
    elementQueue.push({ pElement, 0 });

    while (!elementQueue.empty() && elements.size() < maxElements) {
        if (stopProcessing.load(std::memory_order_acquire)) return;

        ElementInfo current = elementQueue.front();
        elementQueue.pop();

        if (current.depth >= maxDepth) continue;

        if (current.element) {
            elements.push_back(current.element);
            DebugLog(L"Collected element at depth " + std::to_wstring(current.depth));
        }

        CComPtr<IUIAutomationTreeWalker> pControlWalker;
        HRESULT hr = pAutomation->get_ControlViewWalker(&pControlWalker);
        if (FAILED(hr)) {
            DebugLog(L"Failed to get ControlViewWalker: " + std::to_wstring(hr));
            continue;
        }

        CComPtr<IUIAutomationElement> pChild;
        hr = pControlWalker->GetFirstChildElement(current.element, &pChild);
        if (FAILED(hr)) {
            DebugLog(L"Failed to get first child element: " + std::to_wstring(hr));
            continue;
        }

        int childCount = 0;
        while (SUCCEEDED(hr) && pChild && childCount < maxChildren) {
            elementQueue.push({ pChild, current.depth + 1 });

            CComPtr<IUIAutomationElement> pNextSibling;
            hr = pControlWalker->GetNextSiblingElement(pChild, &pNextSibling);
            if (FAILED(hr)) {
                DebugLog(L"Failed to get next sibling element: " + std::to_wstring(hr));
                break;
            }

            pChild = pNextSibling;
            childCount++;
        }
    }
}

void ProcessCursorPosition(POINT point) {
    DebugLog(L"Processing cursor position: (" + std::to_wstring(point.x) + L", " + std::to_wstring(point.y) + L")");

    CComPtr<IUIAutomationElement> pElement = NULL;
    HRESULT hr = pAutomation->ElementFromPoint(point, &pElement);

    if (SUCCEEDED(hr) && pElement) {
        newElementDetected.store(true, std::memory_order_release);
        DebugLog(L"New element detected.");

        std::this_thread::sleep_for(std::chrono::milliseconds(50));
        newElementDetected.store(false, std::memory_order_release);
        DebugLog(L"newElementDetected flag set to false after sleep.");

        if (IsDifferentElement(pElement)) {
            StopCurrentProcesses();

            {
                DebugLog(L"Attempting to lock mtx in ProcessCursorPosition.");
                std::lock_guard<std::mutex> lock(mtx);
                DebugLog(L"Locked mtx in ProcessCursorPosition.");
                if (pPrevElement && rectangleDrawn) {
                    QueueRectangle(prevRect, RectangleAction::Clear);
                }

                hr = pElement->get_CurrentBoundingRectangle(&prevRect);
                if (SUCCEEDED(hr)) {
                    QueueRectangle(prevRect, RectangleAction::Draw);
                }

                pPrevElement.Release();
                pPrevElement = pElement;
            }

            std::vector<CComPtr<IUIAutomationElement>> elements;
            CollectElementsDFS(pElement, elements, MAX_DEPTH, MAX_CHILDREN, MAX_ELEMENTS);

            std::queue<TextRect> textRectQueue;
            size_t currentTextLength = 0;
            std::unordered_set<std::wstring> processedTexts;

            for (auto& element : elements) {
                ReadElementText(element, textRectQueue, currentTextLength, processedTexts);
            }

            if (textRectQueue.empty()) {
                DebugLog(L"textRectQueue is empty after processing elements.");
            }
            else {
                DebugLog(L"textRectQueue populated with " + std::to_wstring(textRectQueue.size()) + L" elements.");
            }

            if (newElementDetected.load(std::memory_order_acquire)) return;

            std::wstring combinedText;

            {
                DebugLog(L"Attempting to lock speechMtx in WorkerThread.");
                std::lock_guard<std::mutex> lock(speechMtx);
                DebugLog(L"Locked speechMtx in WorkerThread.");
                newElementDetected.store(true, std::memory_order_release);

                DebugLog(L"Speech queue size before clearing in WorkerThread: " + std::to_wstring(speechQueue.size()));
                while (!speechQueue.empty()) {
                    speechQueue.pop();
                }
                DebugLog(L"Cleared speech queue in WorkerThread.");

                while (!textRectQueue.empty()) {
                    const auto& textRectPair = textRectQueue.front();
                    if (!combinedText.empty()) {
                        combinedText += L" ";
                    }
                    combinedText += textRectPair.text;
                    speechQueue.push(textRectPair);
                    textRectQueue.pop();
                }

                DebugLog(L"Speech queue size after transfer: " + std::to_wstring(speechQueue.size()));
            }

            if (!combinedText.empty()) {
                PrintText(combinedText);
            }

            speechCv.notify_all();
            DebugLog(L"Notified speech thread.");
        }
    }
    else {
        DebugLog(L"Failed to retrieve UI element from point: " + std::to_wstring(hr));
    }
}

void ProcessCursorPositionAsync(POINT point) {
    submitTask(ProcessCursorPosition, point);
}

void WorkerThread() {
    while (!stopWorker.load(std::memory_order_acquire)) {
        POINT point;
        {
            std::unique_lock<std::mutex> lock(queueMtx);
            DebugLog(L"Waiting on cv in WorkerThread.");
            cv.wait(lock, [] { return !pointsQueue.empty() || stopWorker.load(std::memory_order_acquire); });
            DebugLog(L"Finished waiting on cv in WorkerThread.");

            if (stopWorker.load(std::memory_order_acquire) && pointsQueue.empty()) {
                DebugLog(L"Stopping worker thread.");
                return;
            }

            point = pointsQueue.front();
            pointsQueue.pop();
        }

        ProcessCursorPositionAsync(point);
    }
}

void CheckCursorMovement() {
    while (!stopWorker.load(std::memory_order_acquire)) {
        std::this_thread::sleep_for(std::chrono::milliseconds(300));

        auto now = std::chrono::steady_clock::now();
        if (cursorMoving.load(std::memory_order_acquire) && std::chrono::duration_cast<std::chrono::milliseconds>(now - lastCursorMoveTime).count() > 150) {
            POINT point;
            GetCursorPos(&point);

            {
                std::lock_guard<std::mutex> lock(queueMtx);
                pointsQueue.push(point);
            }
            cv.notify_one();
            cursorMoving.store(false, std::memory_order_release);
            DebugLog(L"Detected cursor movement to: (" + std::to_wstring(point.x) + L", " + std::to_wstring(point.y) + L")");
        }
    }
}

LRESULT CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode >= 0 && wParam == WM_MOUSEMOVE) {
        cursorMoving.store(true, std::memory_order_release); // Set cursor moving flag
        lastCursorMoveTime = std::chrono::steady_clock::now(); // Capture current time
    }
    // Pass the hook information to the next hook procedure in the current hook chain
    return CallNextHookEx(hMouseHook, nCode, wParam, lParam);
}

int main() {
    SetConsoleBufferSize();

    _CrtSetDbgFlag(_CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF);

    if (_setmode(_fileno(stdout), _O_U16TEXT) == -1) {
        return 1;
    }

    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) {
        DebugLog(L"Failed to initialize COM library: " + std::to_wstring(hr));
        return 1;
    }

    hr = CoCreateInstance(__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pAutomation));
    if (FAILED(hr)) {
        DebugLog(L"Failed to create UI Automation instance: " + std::to_wstring(hr));
        CoUninitialize();
        return 1;
    }

    hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void**)&pVoice);
    if (FAILED(hr)) {
        DebugLog(L"Failed to initialize SAPI: " + std::to_wstring(hr));
        pAutomation.Release();
        CoUninitialize();
        return 1;
    }

    pVoice->SetVolume(100);
    pVoice->SetRate(0);

    hMouseHook = SetWindowsHookEx(WH_MOUSE_LL, MouseProc, NULL, 0);
    if (!hMouseHook) {
        DebugLog(L"Failed to set mouse hook.");
        pVoice.Release();
        pAutomation.Release();
        CoUninitialize();
        return 1;
    }

    submitTask(WorkerThread); // Submits the WorkerThread to the thread pool
    submitTask(CheckCursorMovement); // Submits the CheckCursorMovement to the thread pool
    submitTask(RectangleWorker); // Submits the RectangleWorker to the thread pool
    StartSpeechThread(); // Starts the speech thread

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);

        static auto lastReinitTime = std::chrono::steady_clock::now();
        auto now = std::chrono::steady_clock::now();
        if (std::chrono::duration_cast<std::chrono::minutes>(now - lastReinitTime).count() >= 10) {
            ReinitializeAutomation();
            lastReinitTime = now;
        }
    }

    // Stop worker threads
    {
        std::lock_guard<std::mutex> lock(queueMtx);
        stopWorker.store(true, std::memory_order_release);
    }
    cv.notify_all();

    stopRectWorker.store(true, std::memory_order_release);
    rectQueueCv.notify_all();

    // Stop the speech thread if joinable
    if (speechThread.joinable()) {
        stopSpeechThread.store(true, std::memory_order_release);
        speechThread.join();
    }

    if (hMouseHook) {
        UnhookWindowsHookEx(hMouseHook);
    }

    pVoice.Release();
    pAutomation.Release();
    pPrevElement.Release();

    CoUninitialize();

    DebugLog(L"Application exited.");

    return 0;
}


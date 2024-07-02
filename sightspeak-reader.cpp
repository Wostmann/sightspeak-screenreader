#include <iostream>
#include <UIAutomation.h>
#include <windows.h>
#include <tchar.h>
#include <chrono>
#include <mutex>
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
#include <queue>

template <typename T>
class ConcurrentQueue {
public:
    // Add an element to the queue.
    void push(const T& item) {
        std::lock_guard<std::mutex> lock(mutex_);
        queue_.push(item);
        condition_.notify_one();
    }

    // Get the front element. Wait for an element if the queue is empty.
    void pop(T& item) {
        std::unique_lock<std::mutex> lock(mutex_);
        condition_.wait(lock, [this] { return !queue_.empty(); });
        item = queue_.front();
        queue_.pop();
    }

    // Try to get the front element. Return false if the queue is empty.
    bool try_pop(T& item) {
        std::lock_guard<std::mutex> lock(mutex_);
        if (queue_.empty()) {
            return false;
        }
        item = queue_.front();
        queue_.pop();
        return true;
    }

    // Check if the queue is empty.
    bool empty() const {
        std::lock_guard<std::mutex> lock(mutex_);
        return queue_.empty();
    }

    // Get the size of the queue.
    size_t size() const {
        std::lock_guard<std::mutex> lock(mutex_);
        return queue_.size();
    }

private:
    mutable std::mutex mutex_;
    std::queue<T> queue_;
    std::condition_variable condition_;
};

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

struct TextRect {
    std::wstring text;
    RECT rect;
};

// Global variables
HHOOK hMouseHook; // Mouse hook handle
CComPtr<IUIAutomation> pAutomation = NULL; // UI Automation pointer
CComPtr<IUIAutomationElement> pPrevElement = NULL; // Previous UI element pointer
RECT prevRect = { 0, 0, 0, 0 }; // Previous rectangle
std::chrono::time_point<std::chrono::steady_clock> lastProcessedTime; // Last processed time
std::mutex mtx; // General mutex for synchronizing access

ConcurrentQueue<POINT> pointsQueue; // Concurrent queue for storing cursor points
std::atomic<bool> stopWorker(false); // Atomic boolean to signal worker thread to stop

CComPtr<ISpVoice> pVoice = NULL; // Speech voice pointer
std::mutex pVoiceMtx; // Mutex for pVoice

std::atomic<bool> cursorMoving(false); // Atomic boolean to track cursor movement
std::chrono::time_point<std::chrono::steady_clock> lastCursorMoveTime; // Last cursor move time

std::atomic<bool> stopProcessing(false); // Atomic boolean to signal processing to stop

std::mutex speechMtx; // Mutex for speech queue
std::condition_variable speechCv; // Condition variable for speech thread
ConcurrentQueue<TextRect> speechQueue; // Concurrent queue for speech tasks
std::atomic<bool> newElementDetected(false); // Atomic boolean to detect new UI elements

std::thread speechThread; // Thread for speech processing
std::atomic<bool> speaking(false); // Atomic boolean to track speaking status
std::condition_variable speakingCv; // Condition variable to signal end of speaking

std::atomic<bool> rectangleDrawn(false); // Atomic boolean to track rectangle drawing status

const size_t MAX_TEXT_LENGTH = 10000; // Maximum text length
const int MAX_DEPTH = 5; // Maximum depth for UI element traversal
const int MAX_CHILDREN = 20; // Maximum number of children per UI element
const int MAX_ELEMENTS = 200; // Maximum number of UI elements to process

std::chrono::time_point<std::chrono::steady_clock> lastRectangleDrawTime; // Last rectangle draw time

// Shared state for rectangle management
std::atomic<bool> speechOngoing(false); // Atomic boolean to track ongoing speech
RECT currentRect; // Current rectangle being processed
std::mutex rectMtx; // Mutex for rectangle operations
std::condition_variable rectCv; // Condition variable for rectangle thread

enum class RectangleAction { Draw, Clear }; // Enum for rectangle actions

struct RectangleTask {
    RECT rect; // Rectangle coordinates
    RectangleAction action; // Rectangle action (draw or clear)
};

ConcurrentQueue<RectangleTask> rectQueue; // Concurrent queue for rectangle tasks
std::atomic<bool> stopRectWorker(false); // Atomic boolean to signal rectangle worker thread to stop

// Worker threads for various tasks
std::thread worker;              // General worker thread
std::thread rectangleWorker;     // Thread for handling rectangle drawing and clearing
std::thread cursorChecker;       // Thread for checking cursor position

// Logging function to output debug messages to both debug console and log file
void DebugLog(const std::wstring& message) {
    std::wstringstream ws;
    ws << L"[DEBUG] " << message << std::endl;
    OutputDebugString(ws.str().c_str());
    std::wofstream logFile("debug.log", std::ios::app);
    logFile << ws.str();
}

// Set the console buffer size for better visibility of console output
void SetConsoleBufferSize() {
    CONSOLE_SCREEN_BUFFER_INFOEX csbiex = { 0 };
    csbiex.cbSize = sizeof(csbiex);
    HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
    if (GetConsoleScreenBufferInfoEx(hConsole, &csbiex)) {
        csbiex.dwSize.X = 8000; // Set the width of the buffer
        csbiex.dwSize.Y = 2000; // Set the height of the buffer
        SetConsoleScreenBufferInfoEx(hConsole, &csbiex);
    }
}

// Queue a rectangle task for drawing or clearing
void QueueRectangle(const RECT& rect, RectangleAction action) {
    rectQueue.push({ rect, action });  // Push the task to the queue
    rectCv.notify_one();               // Notify the rectangle worker
    DebugLog((action == RectangleAction::Draw ? L"Queueing rectangle for drawing: (" : L"Queueing rectangle for clearing: (") +
        std::to_wstring(rect.left) + L", " +
        std::to_wstring(rect.top) + L", " +
        std::to_wstring(rect.right) + L", " +
        std::to_wstring(rect.bottom) + L")");
}

// Process a rectangle task (draw or clear)
void ProcessRectangle(const RECT& rect, RectangleAction action) {
    HDC hdc = GetDC(NULL); // Get device context
    if (hdc) {
        if (action == RectangleAction::Draw) {
            // Create and select a red pen for drawing
            HPEN hPen = CreatePen(PS_SOLID, 2, RGB(255, 0, 0));
            HGDIOBJ hOldPen = SelectObject(hdc, hPen);
            HBRUSH hBrush = (HBRUSH)GetStockObject(HOLLOW_BRUSH);
            HGDIOBJ hOldBrush = SelectObject(hdc, hBrush);

            // Draw the rectangle
            Rectangle(hdc, rect.left, rect.top, rect.right, rect.bottom);

            // Restore old objects and delete the created pen
            SelectObject(hdc, hOldPen);
            SelectObject(hdc, hOldBrush);
            DeleteObject(hPen);
        }
        else {
            // Invalidate the rectangle area to force a redraw
            InvalidateRect(NULL, &rect, TRUE);
            DebugLog(L"Invalidated rectangle area: (" +
                std::to_wstring(rect.left) + L", " +
                std::to_wstring(rect.top) + L", " +
                std::to_wstring(rect.right) + L", " +
                std::to_wstring(rect.bottom) + L")");
        }
        ReleaseDC(NULL, hdc); // Release device context
    }
    else {
        DebugLog(L"Failed to get device context.");
    }
}

// Worker thread function to handle rectangle tasks
void RectangleWorker() {
    DebugLog(L"Starting RectangleWorker.");
    while (!stopRectWorker.load()) {
        RectangleTask task;
        if (rectQueue.try_pop(task)) {
            DebugLog(L"Processing rectangle: (" +
                std::to_wstring(task.rect.left) + L", " +
                std::to_wstring(task.rect.top) + L", " +
                std::to_wstring(task.rect.right) + L", " +
                std::to_wstring(task.rect.bottom) + L")");
            ProcessRectangle(task.rect, task.action);
        }
        else {
            std::unique_lock<std::mutex> lock(rectMtx);
            rectCv.wait(lock, [] { return !rectQueue.empty() || stopRectWorker.load(); });
        }
    }
    DebugLog(L"Exiting RectangleWorker.");
}

// Print text to console and log it
void PrintText(const std::wstring& text) {
    std::wcout << text << std::endl;
    std::wcout.flush();
    DebugLog(L"Printed text to console: " + text);
}

// Stop speech and clear the speech queue
void StopSpeechAndClearQueue() {
    DebugLog(L"Entering StopSpeechAndClearQueue.");
    std::scoped_lock lock(pVoiceMtx, speechMtx); // Lock both mutexes at once to avoid deadlock

    // Stop ongoing speech
    if (pVoice) {
        HRESULT hr = pVoice->Speak(nullptr, SPF_PURGEBEFORESPEAK, nullptr);
        if (FAILED(hr)) {
            DebugLog(L"Failed to stop speech: " + std::to_wstring(hr));
        }
        else {
            DebugLog(L"Successfully stopped speech.");
        }
    }

    // Clear the speech queue
    TextRect temp;
    while (speechQueue.try_pop(temp)) {
        // Nothing to do here, just clearing the queue
    }

    // Wait until the speaking flag is false
    std::unique_lock<std::mutex> ul(speechMtx);
    speakingCv.wait(ul, [] { return !speaking.load(); });

    DebugLog(L"Exiting StopSpeechAndClearQueue.");
}

// Stop all current processes
void StopCurrentProcesses() {
    DebugLog(L"Entering StopCurrentProcesses.");
    try {
        // Stop ongoing speech
        {
            std::scoped_lock lock(pVoiceMtx, speechMtx); // Lock both mutexes at once to avoid deadlock
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
        TextRect temp;
        while (speechQueue.try_pop(temp)) {}

        // Ensure speaking flag is false
        {
            std::unique_lock<std::mutex> ul(speechMtx);
            speakingCv.wait(ul, [] { return !speaking.load(); });
        }

        // Clear the current rectangle
        RECT tempRect;
        {
            std::lock_guard<std::mutex> rectLock(rectMtx);
            tempRect = currentRect;
            DebugLog(L"Current rectangle to be cleared: (" + std::to_wstring(tempRect.left) + L", " +
                std::to_wstring(tempRect.top) + L", " +
                std::to_wstring(tempRect.right) + L", " +
                std::to_wstring(tempRect.bottom) + L")");
        }
        QueueRectangle(tempRect, RectangleAction::Clear);
    }
    catch (const std::system_error& e) {
        DebugLog(L"Exception in StopCurrentProcesses: " + Utf8ToWstring(e.what()));
    }
    catch (const std::exception& e) {
        DebugLog(L"Exception in StopCurrentProcesses: " + Utf8ToWstring(e.what()));
    }
    DebugLog(L"Exiting StopCurrentProcesses.");
}

// Speech thread function
void SpeakText() {
    DebugLog(L"SpeakText thread started.");
    try {
        while (!stopWorker.load()) {
            TextRect textRectPair;
            {
                std::unique_lock<std::mutex> lock(speechMtx);
                speechCv.wait(lock, [] { return !speechQueue.empty() || stopWorker.load(); }); // Wait for text in the queue or stop signal
                if (stopWorker.load()) break; // Exit if stop signal is received
                if (speechQueue.try_pop(textRectPair)) {
                    // Process textRectPair
                }
            }

            const std::wstring& textToSpeak = textRectPair.text;
            const RECT& textRect = textRectPair.rect;

            if (textToSpeak.empty()) {
                DebugLog(L"Text to speak is empty, skipping.");
                continue;
            }

            DebugLog(L"Attempting to speak text: " + textToSpeak);
            speaking.store(true); // Set speaking flag to true

            {
                std::lock_guard<std::mutex> lock(pVoiceMtx);
                if (pVoice) {
                    HRESULT hr = pVoice->Speak(nullptr, SPF_PURGEBEFORESPEAK, nullptr); // Stop any ongoing speech
                    if (FAILED(hr)) {
                        DebugLog(L"Failed to purge speech: " + std::to_wstring(hr));
                        speaking.store(false);
                        continue;
                    }

                    hr = pVoice->Speak(textToSpeak.c_str(), SPF_ASYNC, nullptr); // Speak the text asynchronously
                    if (FAILED(hr)) {
                        DebugLog(L"Failed to speak text: " + textToSpeak + L" Error: " + std::to_wstring(hr));
                        speaking.store(false);
                        continue;
                    }
                }
            }

            {
                std::lock_guard<std::mutex> rectLock(rectMtx);
                currentRect = textRect; // Set the current rectangle
            }
            QueueRectangle(currentRect, RectangleAction::Draw); // Queue rectangle drawing

            SPVOICESTATUS status;
            while (speaking.load()) {
                {
                    std::lock_guard<std::mutex> lock(pVoiceMtx);
                    if (pVoice) {
                        HRESULT hr = pVoice->GetStatus(&status, nullptr); // Get speech status
                        if (FAILED(hr)) {
                            DebugLog(L"Failed to get status: " + std::to_wstring(hr));
                            break;
                        }

                        if (status.dwRunningState == SPRS_DONE) {
                            DebugLog(L"Speech done for text: " + textToSpeak);
                            break;
                        }
                    }
                }

                std::this_thread::sleep_for(std::chrono::milliseconds(100)); // Sleep for a while before checking status again
            }

            QueueRectangle(currentRect, RectangleAction::Clear); // Queue rectangle clearing

            speaking.store(false); // Set speaking flag to false
            speakingCv.notify_all(); // Notify that speaking has finished
            DebugLog(L"Finished processing text: " + textToSpeak);
        }
    }
    catch (const std::system_error& e) {
        DebugLog(L"Exception in SpeakText: " + Utf8ToWstring(e.what()));
    }
    catch (const std::exception& e) {
        DebugLog(L"Exception in SpeakText: " + Utf8ToWstring(e.what()));
    }
}

// Start the speech thread
void StartSpeechThread() {
    stopWorker.store(false); // Reset stop signal
    speechThread = std::thread(SpeakText); // Start the speech thread
    DebugLog(L"Started speech thread.");
}

// Stop the speech thread
void StopSpeechThread() {
    stopWorker.store(true); // Set stop signal
    speechCv.notify_all(); // Notify the speech thread
    if (speechThread.joinable()) {
        speechThread.join(); // Join the speech thread
    }
    DebugLog(L"Stopped speech thread.");
}

// Check if the UI element is different from the previous one
bool IsDifferentElement(CComPtr<IUIAutomationElement> pElement) {
    DebugLog(L"Attempting to lock mtx in IsDifferentElement.");
    std::lock_guard<std::mutex> lock(mtx); // Lock mutex for thread safety
    DebugLog(L"Locked mtx in IsDifferentElement.");
    if (pPrevElement == NULL && pElement == NULL) {
        return false;
    }
    else if (pPrevElement == NULL || pElement == NULL) {
        return true;
    }

    BOOL areSame;
    HRESULT hr = pAutomation->CompareElements(pPrevElement, pElement, &areSame); // Compare current and previous elements
    DebugLog(L"Compared elements in IsDifferentElement.");
    return SUCCEEDED(hr) && !areSame; // Return true if elements are different
}

// Reinitialize UI Automation and COM components
void ReinitializeAutomation() {
    DebugLog(L"Reinitializing UI Automation and COM components");

    pAutomation.Release();
    pPrevElement.Release();
    {
        std::lock_guard<std::mutex> lock(pVoiceMtx);
        pVoice.Release();
    }

    HRESULT hr = CoInitialize(NULL); // Initialize COM library
    if (FAILED(hr)) {
        DebugLog(L"Failed to reinitialize COM library: " + std::to_wstring(hr));
        return;
    }

    hr = CoCreateInstance(__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pAutomation)); // Create UI Automation instance
    if (FAILED(hr)) {
        DebugLog(L"Failed to reinitialize UI Automation: " + std::to_wstring(hr));
        CoUninitialize();
        return;
    }

    hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void**)&pVoice); // Create SAPI voice instance
    if (FAILED(hr)) {
        DebugLog(L"Failed to reinitialize SAPI: " + std::to_wstring(hr));
        pAutomation.Release();
        CoUninitialize();
        return;
    }

    {
        std::lock_guard<std::mutex> lock(pVoiceMtx);
        pVoice->SetVolume(100); // Set voice volume
        pVoice->SetRate(0); // Set voice rate
    }

    DebugLog(L"Successfully reinitialized UI Automation and COM components");
}

// Read text from a UI element
void ReadElementText(CComPtr<IUIAutomationElement> pElement, ConcurrentQueue<TextRect>& textRectQueue, size_t& currentTextLength, std::unordered_set<std::wstring>& processedTexts) {
    if (stopProcessing.load()) return; // Exit if processing is stopped

    try {
        CComPtr<IUIAutomationTextPattern> pTextPattern = NULL;
        HRESULT hr = pElement->GetCurrentPatternAs(UIA_TextPatternId, IID_PPV_ARGS(&pTextPattern)); // Get TextPattern from element
        if (SUCCEEDED(hr) && pTextPattern) {
            std::this_thread::sleep_for(std::chrono::milliseconds(100)); // Delay to ensure the text is ready for extraction

            CComPtr<IUIAutomationTextRange> pTextRange = NULL;
            hr = pTextPattern->get_DocumentRange(&pTextRange); // Get document range from TextPattern
            if (SUCCEEDED(hr) && pTextRange) {
                CComBSTR text;
                hr = pTextRange->GetText(-1, &text); // Get the text from the text range
                if (SUCCEEDED(hr)) {
                    std::wstring textStr(static_cast<wchar_t*>(text));
                    if (!textStr.empty() && processedTexts.find(textStr) == processedTexts.end()) {
                        processedTexts.insert(textStr); // Add text to processed set
                        currentTextLength += textStr.length();
                        if (currentTextLength <= MAX_TEXT_LENGTH) {
                            RECT rect = {};
                            SAFEARRAY* pRects = NULL;
                            hr = pTextRange->GetBoundingRectangles(&pRects); // Get bounding rectangles of the text
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
                                SafeArrayDestroy(pRects); // Destroy the safe array
                            }
                            textRectQueue.push({ textStr, rect }); // Add text and its rectangle to the queue
                            DebugLog(L"Text and bounding rectangle pushed to queue.");
                        }
                    }
                }
                else {
                    DebugLog(L"Failed to get text from text range: " + std::to_wstring(hr));
                }
            }
            else {
                DebugLog(L"Failed to get document range: " + std::to_wstring(hr));
            }
        }

        CComBSTR name;
        hr = pElement->get_CurrentName(&name); // Get the name of the element
        if (SUCCEEDED(hr)) {
            std::wstring nameStr(static_cast<wchar_t*>(name));
            if (!nameStr.empty() && processedTexts.find(nameStr) == processedTexts.end()) {
                processedTexts.insert(nameStr); // Add name to processed set
                currentTextLength += nameStr.length();
                if (currentTextLength <= MAX_TEXT_LENGTH) {
                    RECT rect = {};
                    hr = pElement->get_CurrentBoundingRectangle(&rect); // Get bounding rectangle of the element
                    if (SUCCEEDED(hr)) {
                        textRectQueue.push({ nameStr, rect }); // Add name and its rectangle to the queue
                        DebugLog(L"Element name and bounding rectangle pushed to queue.");
                    }
                }
            }
        }
    }
    catch (const std::exception& e) {
        DebugLog(L"Exception in ReadElementText: " + Utf8ToWstring(e.what()));
    }
}



// Collect UI elements using depth-first search
void CollectElementsDFS(CComPtr<IUIAutomationElement> pElement, std::vector<CComPtr<IUIAutomationElement>>& elements, int maxDepth, int maxChildren, int maxElements) {
    if (!pElement) return; // Exit if the element is null

    struct ElementInfo {
        CComPtr<IUIAutomationElement> element;
        int depth;
    };

    ConcurrentQueue<ElementInfo> elementQueue;
    elementQueue.push({ pElement, 0 }); // Push the initial element with depth 0

    while (!elementQueue.empty() && elements.size() < maxElements) {
        if (stopProcessing.load()) return; // Exit if processing is stopped

        ElementInfo current;
        if (elementQueue.try_pop(current)) {
            if (current.depth >= maxDepth) continue; // Skip if max depth is reached

            if (current.element) {
                elements.push_back(current.element); // Add element to the list
                DebugLog(L"Collected element at depth " + std::to_wstring(current.depth));
            }

            CComPtr<IUIAutomationTreeWalker> pControlWalker;
            HRESULT hr = pAutomation->get_ControlViewWalker(&pControlWalker); // Get the control view walker
            if (FAILED(hr)) {
                DebugLog(L"Failed to get ControlViewWalker: " + std::to_wstring(hr));
                continue;
            }

            CComPtr<IUIAutomationElement> pChild;
            hr = pControlWalker->GetFirstChildElement(current.element, &pChild); // Get the first child element
            if (FAILED(hr)) {
                DebugLog(L"Failed to get first child element: " + std::to_wstring(hr));
                continue;
            }

            int childCount = 0;
            while (SUCCEEDED(hr) && pChild && childCount < maxChildren) {
                elementQueue.push({ pChild, current.depth + 1 }); // Push the child element with incremented depth

                CComPtr<IUIAutomationElement> pNextSibling;
                hr = pControlWalker->GetNextSiblingElement(pChild, &pNextSibling); // Get the next sibling element
                if (FAILED(hr)) {
                    DebugLog(L"Failed to get next sibling element: " + std::to_wstring(hr));
                    break;
                }

                pChild = pNextSibling;
                childCount++; // Increment the child count
            }
        }
    }
}


// Process cursor position and detect UI elements
void ProcessCursorPosition(POINT point) {
    DebugLog(L"Processing cursor position: (" + std::to_wstring(point.x) + L", " + std::to_wstring(point.y) + L")");

    CComPtr<IUIAutomationElement> pElement = NULL;
    HRESULT hr = pAutomation->ElementFromPoint(point, &pElement); // Get the UI element from the cursor position

    if (SUCCEEDED(hr) && pElement) {
        {
            std::lock_guard<std::mutex> lock(mtx);
            newElementDetected.store(true); // Indicate a new element is detected
        }
        DebugLog(L"New element detected.");

        // Consider reducing or removing this sleep
        std::this_thread::sleep_for(std::chrono::milliseconds(10)); // Sleep briefly to stabilize detection (reduced to 10 ms)
        {
            std::lock_guard<std::mutex> lock(mtx);
            newElementDetected.store(false); // Reset the flag after sleep
        }
        DebugLog(L"newElementDetected flag set to false after sleep.");

        if (IsDifferentElement(pElement)) {
            StopCurrentProcesses();  // Ensure all current processes are stopped

            {
                DebugLog(L"Attempting to lock mtx in ProcessCursorPosition.");
                std::lock_guard<std::mutex> lock(mtx);
                DebugLog(L"Locked mtx in ProcessCursorPosition.");
                if (pPrevElement && rectangleDrawn) {
                    QueueRectangle(prevRect, RectangleAction::Clear); // Clear previous rectangle
                }

                hr = pElement->get_CurrentBoundingRectangle(&prevRect); // Get bounding rectangle of new element
                if (SUCCEEDED(hr)) {
                    QueueRectangle(prevRect, RectangleAction::Draw); // Draw new rectangle
                }

                pPrevElement.Release();
                pPrevElement = pElement; // Update previous element
            }

            std::vector<CComPtr<IUIAutomationElement>> elements;
            CollectElementsDFS(pElement, elements, MAX_DEPTH, MAX_CHILDREN, MAX_ELEMENTS); // Collect child elements

            ConcurrentQueue<TextRect> textRectQueue;
            size_t currentTextLength = 0;
            std::unordered_set<std::wstring> processedTexts;

            for (auto& element : elements) {
                ReadElementText(element, textRectQueue, currentTextLength, processedTexts); // Read text from collected elements
            }

            if (textRectQueue.empty()) {
                DebugLog(L"textRectQueue is empty after processing elements.");
            }
            else {
                DebugLog(L"textRectQueue populated with elements.");
            }

            if (newElementDetected.load()) return; // Exit if a new element is detected during processing

            std::wstring combinedText;

            {
                DebugLog(L"Attempting to lock speechMtx in WorkerThread.");
                std::lock_guard<std::mutex> lock(speechMtx);
                DebugLog(L"Locked speechMtx in WorkerThread.");
                newElementDetected.store(true);

                DebugLog(L"Speech queue size before clearing in WorkerThread: ");
                TextRect temp;
                while (speechQueue.try_pop(temp)) {} // Clear speech queue
                DebugLog(L"Cleared speech queue in WorkerThread.");

                while (textRectQueue.try_pop(temp)) {
                    if (!combinedText.empty()) {
                        combinedText += L" ";
                    }
                    combinedText += temp.text; // Combine text from textRectQueue
                    speechQueue.push(temp); // Transfer textRect to speechQueue
                }

                DebugLog(L"Speech queue size after transfer: ");
            }

            if (!combinedText.empty()) {
                PrintText(combinedText); // Print combined text to console
            }

            speechCv.notify_all(); // Notify speech thread
            DebugLog(L"Notified speech thread.");
        }
    }
    else {
        DebugLog(L"Failed to retrieve UI element from point: " + std::to_wstring(hr));
    }
}


// Worker thread function to process cursor movements
void WorkerThread() {
    while (!stopWorker.load()) {
        POINT point;
        if (pointsQueue.try_pop(point)) {
            ProcessCursorPosition(point); // Process the cursor position if available
        }
        else {
            std::this_thread::sleep_for(std::chrono::milliseconds(100)); // Sleep if no cursor position to process
        }
    }
}

// Thread function to check cursor movement
void CheckCursorMovement() {
    while (!stopWorker.load()) {
        std::this_thread::sleep_for(std::chrono::milliseconds(50)); // Sleep briefly before checking cursor movement

        auto now = std::chrono::steady_clock::now();
        if (cursorMoving.load() && std::chrono::duration_cast<std::chrono::milliseconds>(now - lastCursorMoveTime).count() > 150) {
            POINT point;
            GetCursorPos(&point); // Get current cursor position

            pointsQueue.push(point); // Push cursor position to the queue
            cursorMoving.store(false); // Reset cursor moving flag
            DebugLog(L"Detected cursor movement to: (" + std::to_wstring(point.x) + L", " + std::to_wstring(point.y) + L")");
        }
    }
}

// Mouse hook procedure to detect mouse movements
LRESULT CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode >= 0 && wParam == WM_MOUSEMOVE) {
        cursorMoving.store(true); // Set cursor moving flag to true
        lastCursorMoveTime = std::chrono::steady_clock::now(); // Update the last cursor move time
    }
    return CallNextHookEx(hMouseHook, nCode, wParam, lParam); // Call the next hook in the chain
}

// Shutdown function to clean up resources
void Shutdown() {
    stopWorker.store(true); // Signal all threads to stop
    speechCv.notify_all(); // Notify all waiting on speech condition variable
    rectCv.notify_all(); // Notify all waiting on rectangle condition variable

    if (speechThread.joinable()) {
        speechThread.join(); // Join speech thread
    }

    if (worker.joinable()) {
        worker.join(); // Join worker thread
    }

    if (rectangleWorker.joinable()) {
        rectangleWorker.join(); // Join rectangle worker thread
    }

    if (cursorChecker.joinable()) {
        cursorChecker.join(); // Join cursor checker thread
    }

    UnhookWindowsHookEx(hMouseHook); // Unhook the mouse hook

    std::lock_guard<std::mutex> lock(pVoiceMtx);
    pVoice.Release(); // Release SAPI voice

    pAutomation.Release(); // Release UI Automation
    pPrevElement.Release(); // Release previous element

    CoUninitialize(); // Uninitialize COM library

    DebugLog(L"Application exited."); // Log application exit
}

int main() {
    SetConsoleBufferSize(); // Set console buffer size for better visibility

    _CrtSetDbgFlag(_CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF); // Enable memory leak checking

    if (_setmode(_fileno(stdout), _O_U16TEXT) == -1) {
        return 1; // Set console mode to Unicode
    }

    HRESULT hr = CoInitialize(NULL); // Initialize COM library
    if (FAILED(hr)) {
        DebugLog(L"Failed to initialize COM library: " + std::to_wstring(hr));
        return 1;
    }

    hr = CoCreateInstance(__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pAutomation)); // Create UI Automation instance
    if (FAILED(hr)) {
        DebugLog(L"Failed to create UI Automation instance: " + std::to_wstring(hr));
        CoUninitialize();
        return 1;
    }

    {
        std::lock_guard<std::mutex> lock(pVoiceMtx);
        hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void**)&pVoice); // Create SAPI voice instance
        if (FAILED(hr)) {
            DebugLog(L"Failed to initialize SAPI: " + std::to_wstring(hr));
            pAutomation.Release();
            CoUninitialize();
            return 1;
        }

        pVoice->SetVolume(100); // Set voice volume
        pVoice->SetRate(0); // Set voice rate
    }

    hMouseHook = SetWindowsHookEx(WH_MOUSE_LL, MouseProc, NULL, 0); // Set mouse hook
    if (!hMouseHook) {
        DebugLog(L"Failed to set mouse hook.");
        Shutdown(); // Shutdown if mouse hook fails
        return 1;
    }

    worker = std::thread(WorkerThread); // Start worker thread
    cursorChecker = std::thread(CheckCursorMovement); // Start cursor checker thread
    rectangleWorker = std::thread(RectangleWorker); // Start rectangle worker thread
    StartSpeechThread(); // Start speech thread

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) { // Message loop
        TranslateMessage(&msg);
        DispatchMessage(&msg);

        static auto lastReinitTime = std::chrono::steady_clock::now();
        auto now = std::chrono::steady_clock::now();
        if (std::chrono::duration_cast<std::chrono::minutes>(now - lastReinitTime).count() >= 10) {
            ReinitializeAutomation(); // Reinitialize automation every 10 minutes
            lastReinitTime = now;
        }
    }

    Shutdown(); // Perform cleanup on exit

    return 0;
}


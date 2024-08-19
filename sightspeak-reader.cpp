#include <windows.h>
#include <WinSock2.h>
#include <Ws2tcpip.h>
#include <atlbase.h>
#include <UIAutomation.h>
#include <tchar.h>
#include <chrono>
#include <mutex>
#include <thread>
#include <condition_variable>
#include <crtdbg.h>
#include <fcntl.h>
#include <io.h>
#include <algorithm>
#include <sapi.h>
#include <atomic>
#include <unordered_set>
#include <iostream>
#include <sstream>
#include <fstream>
#include <queue>
#include <functional>
#include <tbb/tbb.h>
#include <tbb/concurrent_queue.h>
#include <shared_mutex>
#include <future>
#include <boost/asio.hpp>
#include <boost/thread/thread.hpp>
#include <boost/asio/steady_timer.hpp>

// Utility function to convert UTF-8 string to wide string
std::wstring Utf8ToWstring(const std::string& str) {
    if (str.empty()) return std::wstring();
    int size_needed = MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), NULL, 0);
    std::wstring wstrTo(size_needed, 0);
    MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), &wstrTo[0], size_needed);
    return wstrTo;
}

// Utility function to convert wide string to UTF-8 string
std::string WstringToUtf8(const std::wstring& wstr) {
    if (wstr.empty()) return std::string();
    int size_needed = WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), NULL, 0, NULL, NULL);
    std::string strTo(size_needed, 0);
    WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), &strTo[0], size_needed, NULL, NULL);
    return strTo;
}

// Structure to hold text and its associated rectangle
struct TextRect {
    std::wstring text;
    RECT rect{ 0, 0, 0, 0 };
};

// Global variables for UI Automation and speech synthesis
HHOOK hMouseHook;
CComPtr<IUIAutomation> pAutomation = NULL;
CComPtr<IUIAutomationElement> pPrevElement = NULL;
RECT prevRect = { 0, 0, 0, 0 };
std::shared_mutex mtx;
CComPtr<ISpVoice> pVoice = NULL;
std::mutex pVoiceMtx;
std::atomic<bool> speaking(false);
std::atomic<bool> rectangleDrawn(false);
const size_t MAX_TEXT_LENGTH = 10000;
const int MAX_DEPTH = 5;
const int MAX_CHILDREN = 20;
const int MAX_ELEMENTS = 200;
RECT currentRect;
std::mutex rectMtx;
std::atomic<bool> capsLockOverride(false);
std::chrono::time_point<std::chrono::steady_clock> lastCapsLockPress;
std::shared_mutex elementMutex;
boost::asio::io_context io_context;
std::atomic<std::chrono::time_point<std::chrono::steady_clock>> lastProcessed;
POINT lastProcessedPoint = { -1, -1 };
std::unordered_set<std::wstring> processedTexts;

std::mutex cancelMtx;
std::promise<void> cancelPromise;
std::shared_future<void> cancelFuture = cancelPromise.get_future().share();



// Logging function to output debug messages to both debug console and log file
void DebugLog(const std::wstring& message) {
    std::wstringstream ws;
    ws << L"[DEBUG] " << message << std::endl;
    OutputDebugString(ws.str().c_str());
    std::wofstream logFile("debug.log", std::ios::app);
    logFile << ws.str();
}

// Function to toggle CAPSLOCK override state
void ToggleCapsLockOverride() {
    capsLockOverride.store(!capsLockOverride.load());
    DebugLog(L"CAPSLOCK override " + std::wstring(capsLockOverride.load() ? L"enabled" : L"disabled"));
}

// Forward declaration of ProcessNewElement function
void ProcessNewElement(CComPtr<IUIAutomationElement> pElement);

// Function to move to the parent element
void MoveToParentElement() {
    std::shared_lock<std::shared_mutex> lock(elementMutex);  // Use shared_lock for read-only access
    if (!pPrevElement) return;

    CComPtr<IUIAutomationTreeWalker> pControlWalker;
    HRESULT hr = pAutomation->get_ControlViewWalker(&pControlWalker);
    if (SUCCEEDED(hr)) {
        CComPtr<IUIAutomationElement> pParent;
        hr = pControlWalker->GetParentElement(pPrevElement, &pParent);
        if (SUCCEEDED(hr) && pParent) {
            pPrevElement = pParent;
            DebugLog(L"Moved to parent element");
            ProcessNewElement(pPrevElement);
        }
    }
}

// Function to move to the first child element
void MoveToFirstChildElement() {
    std::shared_lock<std::shared_mutex> lock(elementMutex);  // Use shared_lock for read-only access
    if (!pPrevElement) return;

    CComPtr<IUIAutomationTreeWalker> pControlWalker;
    HRESULT hr = pAutomation->get_ControlViewWalker(&pControlWalker);
    if (SUCCEEDED(hr)) {
        CComPtr<IUIAutomationElement> pChild;
        hr = pControlWalker->GetFirstChildElement(pPrevElement, &pChild);
        if (SUCCEEDED(hr) && pChild) {
            pPrevElement = pChild;
            DebugLog(L"Moved to first child element");
            ProcessNewElement(pPrevElement);
        }
    }
}

// Function to move to the next sibling element
void MoveToNextSiblingElement() {
    std::shared_lock<std::shared_mutex> lock(elementMutex);  // Use shared_lock for read-only access
    if (!pPrevElement) return;

    CComPtr<IUIAutomationTreeWalker> pControlWalker;
    HRESULT hr = pAutomation->get_ControlViewWalker(&pControlWalker);
    if (SUCCEEDED(hr)) {
        CComPtr<IUIAutomationElement> pNextSibling;
        hr = pControlWalker->GetNextSiblingElement(pPrevElement, &pNextSibling);
        if (SUCCEEDED(hr) && pNextSibling) {
            pPrevElement = pNextSibling;
            DebugLog(L"Moved to next sibling element");
            ProcessNewElement(pPrevElement);
        }
    }
}

// Function to move to the previous sibling element
void MoveToPreviousSiblingElement() {
    std::shared_lock<std::shared_mutex> lock(elementMutex);  // Use shared_lock for read-only access
    if (!pPrevElement) return;

    CComPtr<IUIAutomationTreeWalker> pControlWalker;
    HRESULT hr = pAutomation->get_ControlViewWalker(&pControlWalker);
    if (SUCCEEDED(hr)) {
        CComPtr<IUIAutomationElement> pPreviousSibling;
        hr = pControlWalker->GetPreviousSiblingElement(pPrevElement, &pPreviousSibling);
        if (SUCCEEDED(hr) && pPreviousSibling) {
            pPrevElement = pPreviousSibling;
            DebugLog(L"Moved to previous sibling element");
            ProcessNewElement(pPrevElement);
        }
    }
}

// Function to redo the current element
void RedoCurrentElement() {
    std::shared_lock<std::shared_mutex> lock(elementMutex);  // Use shared_lock for read-only access
    if (pPrevElement) {
        DebugLog(L"Redoing current element");
        ProcessNewElement(pPrevElement);
    }
}

// Low-level keyboard procedure to handle keyboard shortcuts
LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode == HC_ACTION) {
        KBDLLHOOKSTRUCT* pKeyBoard = (KBDLLHOOKSTRUCT*)lParam;
        if (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN) {
            if (pKeyBoard->vkCode == VK_CAPITAL) {
                auto now = std::chrono::steady_clock::now();
                if (std::chrono::duration_cast<std::chrono::milliseconds>(now - lastCapsLockPress).count() < 500) {
                    ToggleCapsLockOverride();
                    DebugLog(L"Double CAPSLOCK detected, toggling override.");
                }
                else {
                    DebugLog(L"Single CAPSLOCK press detected.");
                }
                lastCapsLockPress = now;
            }

            if (capsLockOverride.load()) {
                switch (pKeyBoard->vkCode) {
                case 'W':
                    MoveToParentElement();
                    DebugLog(L"CAPSLOCK+W: Move to parent element.");
                    break;
                case 'S':
                    MoveToFirstChildElement();
                    DebugLog(L"CAPSLOCK+S: Move to first child element.");
                    break;
                case 'D':
                    MoveToNextSiblingElement();
                    DebugLog(L"CAPSLOCK+D: Move to next sibling element.");
                    break;
                case 'A':
                    MoveToPreviousSiblingElement();
                    DebugLog(L"CAPSLOCK+A: Move to previous sibling element.");
                    break;
                case 'E':
                    RedoCurrentElement();
                    DebugLog(L"CAPSLOCK+E: Redo current element.");
                    break;
                }
            }
        }
    }
    return CallNextHookEx(NULL, nCode, wParam, lParam);
}

// Function to set the low-level keyboard hook
void SetLowLevelKeyboardHook() {
    DebugLog(L"Setting low-level keyboard hook.");

    HHOOK hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, LowLevelKeyboardProc, NULL, 0);
    if (!hKeyboardHook) {
        DebugLog(L"Failed to set keyboard hook. Error: " + std::to_wstring(GetLastError()));
        return;
    }

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    DebugLog(L"Unhooking keyboard hook.");
    UnhookWindowsHookEx(hKeyboardHook);
}

// Function to set the console buffer size for better visibility of console output
void SetConsoleBufferSize() {
    CONSOLE_SCREEN_BUFFER_INFOEX csbiex = { 0 };
    csbiex.cbSize = sizeof(csbiex);
    HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
    if (GetConsoleScreenBufferInfoEx(hConsole, &csbiex)) {
        csbiex.dwSize.X = 8000;
        csbiex.dwSize.Y = 2000;
        SetConsoleScreenBufferInfoEx(hConsole, &csbiex);
    }
}


void ProcessRectangle(const RECT& rect, bool draw, std::shared_future<void> cancelFuture) {
    if (cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) {
        return;
    }

    auto start = std::chrono::steady_clock::now();

    try {
        DebugLog(L"ProcessRectangle with action: " + std::to_wstring(draw) + L" for rectangle: (" +
            std::to_wstring(rect.left) + L", " +
            std::to_wstring(rect.top) + L", " +
            std::to_wstring(rect.right) + L", " +
            std::to_wstring(rect.bottom) + L")");
        HDC hdc = GetDC(NULL);
        if (hdc) {
            if (draw) {
                
                // Delay handling is managed outside the critical section to avoid locking overhead
                std::this_thread::sleep_for(std::chrono::milliseconds(30));
                HPEN hPen = CreatePen(PS_SOLID, 2, GetSysColor(COLOR_HIGHLIGHT)); // System color that adapts to current theme;
                HGDIOBJ hOldPen = SelectObject(hdc, hPen);
                HBRUSH hBrush = (HBRUSH)GetStockObject(HOLLOW_BRUSH);
                HGDIOBJ hOldBrush = SelectObject(hdc, hBrush);

                Rectangle(hdc, rect.left, rect.top, rect.right, rect.bottom);

                SelectObject(hdc, hOldPen);
                SelectObject(hdc, hOldBrush);
                DeleteObject(hPen);
                rectangleDrawn.store(true);
                {
                    // Lock the rectangle resource to ensure thread-safe access
                    std::lock_guard<std::mutex> rectLock(rectMtx);
                    currentRect = rect; // Update the current rectangle
                }
            }
            else {
                InvalidateRect(NULL, &rect, TRUE);
                rectangleDrawn.store(false);
            }
            ReleaseDC(NULL, hdc);
        }
        else {
            DebugLog(L"Failed to get device context.");
        }
    }
    catch (const std::exception& e) {
        DebugLog(L"Exception in ProcessRectangle: " + Utf8ToWstring(e.what()));
    }
    auto end = std::chrono::steady_clock::now();
    std::chrono::duration<double> elapsed_seconds = end - start;
    DebugLog(L"Processed rectangle " + std::to_wstring(draw) + L" in " + std::to_wstring(elapsed_seconds.count()) + L" seconds.");
}

// Function to print text to the console and log it
void PrintText(const std::wstring& text) {
    std::wcout << text << std::endl;
    std::wcout.flush();
    DebugLog(L"Printed text to console: " + text);
}

// Task to speak text and manage rectangle
std::future<void> SpeakTextTask(const std::wstring& textToSpeak, std::shared_future<void> cancelFuture) {
    return std::async(std::launch::async, [=]() {
        // Check if the task should be canceled before starting
        if (cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) {
            return; // Exit if cancellation is requested
        }
        DebugLog(L"SpeakTextTask started.");

        try {
            // If the text is empty or cancellation is requested, exit early
            if (textToSpeak.empty() || cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) {
                DebugLog(L"Task cancelled before speaking.");
                return;
            }

            PrintText(textToSpeak); // Output the text to the console and log it

            speaking.store(true); // Set the speaking flag to true, indicating speech is in progress

            {
                // Lock the speech synthesis resource to ensure thread-safe access
                std::lock_guard<std::mutex> lock(pVoiceMtx);
                if (pVoice) { // Check if the speech synthesis object is valid
                    // Check for cancellation again before starting speech
                    if (cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) {
                        speaking.store(false); // Reset the speaking flag if canceled
                        return;
                    }

                    // Purge any previous speech to start fresh
                    HRESULT hr = pVoice->Speak(nullptr, SPF_PURGEBEFORESPEAK, nullptr);
                    if (FAILED(hr)) {
                        DebugLog(L"Failed to purge speech: " + std::to_wstring(hr));
                        speaking.store(false); // Reset the speaking flag if purging failed
                        return;
                    }

                    // Start speaking the text asynchronously
                    hr = pVoice->Speak(textToSpeak.c_str(), SPF_ASYNC, nullptr);
                    if (FAILED(hr)) {
                        DebugLog(L"Failed to speak text: " + textToSpeak + L" Error: " + std::to_wstring(hr));
                        speaking.store(false); // Reset the speaking flag if speaking failed
                        return;
                    }
                }
            }

            SPVOICESTATUS status; // Structure to hold the status of the speech synthesis
            while (true) {
                // Check for cancellation every 10 milliseconds
                if (cancelFuture.wait_for(std::chrono::milliseconds(10)) == std::future_status::ready) {
                    DebugLog(L"Task cancelled during speech.");
                    break; // Exit the loop if cancellation is requested
                }

                // Lock the speech synthesis resource to check its status
                std::lock_guard<std::mutex> lock(pVoiceMtx);
                if (pVoice) { // Check if the speech synthesis object is valid
                    HRESULT hr = pVoice->GetStatus(&status, nullptr); // Get the current status of the speech synthesis
                    if (FAILED(hr)) {
                        DebugLog(L"Failed to get status: " + std::to_wstring(hr));
                        break; // Exit the loop if getting the status failed
                    }

                    // Check if the speech synthesis has completed
                    if (status.dwRunningState == SPRS_DONE) {
                        DebugLog(L"Speech done for text: " + textToSpeak);
                        break; // Exit the loop if the speech is done
                    }
                }
            }

            speaking.store(false); // Reset the speaking flag to indicate speech is complete
            DebugLog(L"Finished processing text: " + textToSpeak); // Log the completion of the task
        }
        catch (const std::exception& e) {
            // Log any exceptions that occur during the task
            DebugLog(L"Exception in SpeakTextTask: " + Utf8ToWstring(e.what()));
        }
        });
}

// Class to manage the queue for processing TextRect objects
class ProcessTextRectQueue {
public:
    static void Enqueue(TextRect textRect, std::shared_future<void> cancelFuture) {
        if (cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) {
            return;
        }
        {
            std::lock_guard<std::mutex> lock(queueMutex);
            textRectQueue.push(textRect);
        }
        if(!speaking.load()){
            io_context.post(std::bind(&ProcessTextRectQueue::DequeueAndProcess, cancelFuture));
        }
    }

    static void DequeueAndProcess(std::shared_future<void> cancelFuture) {
        if (textRectQueue.empty() || cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) {
            speaking.store(false);
            return;
        }
        
        speaking.store(true);
        TextRect textRect;
        {
            std::lock_guard<std::mutex> lock(queueMutex);
            textRect = textRectQueue.front();
            textRectQueue.pop();
        }
        DebugLog(L"Dequeued " + textRect.text);
        // Process the rectangle drawing first
        std::async(std::launch::async, ProcessRectangle, textRect.rect, true, cancelFuture);
        

        // Handle the text and rectangle in a non-blocking way
        std::future<void> future = SpeakTextTask(textRect.text, cancelFuture);
        future.get();

        // Clear the rectangle after speaking is done
        ProcessRectangle(textRect.rect, false, cancelFuture);

        // Reset the speaking flag after task completion
        io_context.post(std::bind(&ProcessTextRectQueue::DequeueAndProcess, cancelFuture));
    }

    static void ClearQueue() {
        std::lock_guard<std::mutex> lock(queueMutex);
        std::queue<TextRect> empty;
        std::swap(textRectQueue, empty);
        speaking.store(false);  // Ensure that speaking flag is reset
    }

private:
    static std::queue<TextRect> textRectQueue;
    static std::mutex queueMutex;
    static std::atomic<bool> speaking;
};

std::queue<TextRect> ProcessTextRectQueue::textRectQueue;
std::mutex ProcessTextRectQueue::queueMutex;
std::atomic<bool> ProcessTextRectQueue::speaking = false;

// Function to read text and rectangle from element
void ReadElementText(CComPtr<IUIAutomationElement> pElement, std::shared_future<void> cancelFuture) {
    if (cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) {
        return;  // Exit if cancellation is requested
    }

    try {
        CComPtr<IUIAutomationTextPattern> pTextPattern = NULL;
        HRESULT hr = pElement->GetCurrentPatternAs(UIA_TextPatternId, IID_PPV_ARGS(&pTextPattern));
        if (SUCCEEDED(hr) && pTextPattern) {
            CComPtr<IUIAutomationTextRange> pTextRange = NULL;
            hr = pTextPattern->get_DocumentRange(&pTextRange);
            if (SUCCEEDED(hr) && pTextRange) {
                CComBSTR text;
                hr = pTextRange->GetText(-1, &text);
                if (SUCCEEDED(hr) && text!=NULL) {
                    std::wstring textStr(static_cast<wchar_t*>(text));
                    if (!textStr.empty() && processedTexts.find(textStr) == processedTexts.end()) {
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

                        ProcessTextRectQueue::Enqueue({ textStr, rect }, cancelFuture);
                        processedTexts.insert(textStr);  // Add the text to the set after enqueueing
                        DebugLog(textStr + L" and bounding rectangle pushed to queue.");
                    }
                }
            }
        }

        CComBSTR name;
        hr = pElement->get_CurrentName(&name);
        if (SUCCEEDED(hr) && name != NULL) {
            std::wstring nameStr(static_cast<wchar_t*>(name));
            if (!nameStr.empty() && processedTexts.find(nameStr) == processedTexts.end()) {
                RECT rect = {};
                hr = pElement->get_CurrentBoundingRectangle(&rect);
                if (SUCCEEDED(hr)) {
                    ProcessTextRectQueue::Enqueue({ nameStr, rect }, cancelFuture);
                    processedTexts.insert(nameStr);  // Add the name to the set after enqueueing
                    DebugLog(nameStr + L" and bounding rectangle pushed to queue.");
                }
            }
        }
    }
    catch (const std::exception& e) {
        DebugLog(L"Exception in ReadElementText: " + Utf8ToWstring(e.what()));
    }
}


// Collect UI elements using depth-first search
void CollectElementsBFS(CComPtr<IUIAutomationElement> pElement, std::shared_future<void> cancelFuture) {
    if (!pElement || cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) return;
    processedTexts.clear();
    struct ElementInfo {
        CComPtr<IUIAutomationElement> element;
        int depth{ 0 };
    };

    std::queue<ElementInfo> elementQueue;
    elementQueue.push({ pElement, 0 });

    while (!elementQueue.empty()) {
        if (cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) return;

        ElementInfo current = std::move(elementQueue.front());
        elementQueue.pop();
        if (current.depth >= MAX_DEPTH) continue;

        if (cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) return;
        ReadElementText(current.element, cancelFuture);

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
        while (SUCCEEDED(hr) && pChild && childCount < MAX_CHILDREN) {
            if (cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) return;

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

// Function to stop current processes asynchronously
void StopCurrentProcesses() {
    DebugLog(L"Entering StopCurrentProcesses.");

    try {
        // Signal cancellation of current tasks
        {
            std::lock_guard<std::mutex> cancelLock(cancelMtx);
            cancelPromise.set_value();
            cancelPromise = std::promise<void>();  // Reset for new tasks
            cancelFuture = cancelPromise.get_future().share();
        }

        // Stop ongoing processes
        ProcessTextRectQueue::ClearQueue();
            
        if (rectangleDrawn.load()) {  // Clear the rectangle
            std::lock_guard<std::mutex> rectLock(rectMtx);
            io_context.post([=]() {ProcessRectangle(currentRect, false, cancelFuture);});
        }

        {   //Clear speech
            std::lock_guard<std::mutex> voiceLock(pVoiceMtx);
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

    }
    catch (const std::system_error& e) {
        DebugLog(L"Exception in StopCurrentProcessesAsync: " + Utf8ToWstring(e.what()));
    }
    catch (const std::exception& e) {
        DebugLog(L"Exception in StopCurrentProcessesAsync: " + Utf8ToWstring(e.what()));
    }

    DebugLog(L"Completed StopCurrentProcesses.");
}


// Process new UI element
void ProcessNewElement(CComPtr<IUIAutomationElement> pElement) {
    auto start = std::chrono::steady_clock::now();

    io_context.dispatch([pElement]() {
        StopCurrentProcesses();  // Stop all current tasks

        CollectElementsBFS(pElement, cancelFuture);  // Proceed with BFS to collect UI elements
        DebugLog(L"BFS completed.");
        });

    auto end = std::chrono::steady_clock::now();
    std::chrono::duration<double> elapsed_seconds = end - start;
    DebugLog(L"Processed new element in " + std::to_wstring(elapsed_seconds.count()) + L" seconds.");
}

// Check if the UI element is different from the previous one
bool IsDifferentElement(CComPtr<IUIAutomationElement> pElement) {
    std::shared_lock<std::shared_mutex> lock(mtx);  // Use shared_lock for read-only access
    if (pPrevElement == NULL && pElement == NULL) {
        return false;
    }
    else if (pPrevElement == NULL || pElement == NULL) {
        return true;
    }

    BOOL areSame;
    HRESULT hr = pAutomation->CompareElements(pPrevElement, pElement, &areSame);
    return SUCCEEDED(hr) && !areSame;
}

// Process cursor position and detect UI elements
void ProcessCursorPosition(POINT point) {
    DebugLog(L"Processing cursor position: (" + std::to_wstring(point.x) + L", " + std::to_wstring(point.y) + L")");

    CComPtr<IUIAutomationElement> pElement = NULL;
    HRESULT hr = pAutomation->ElementFromPoint(point, &pElement);

    if (SUCCEEDED(hr) && pElement) {
        if (IsDifferentElement(pElement)) {
            DebugLog(L"New element detected.");
            pPrevElement.Release();
            pPrevElement = pElement; // Update previous element
            ProcessNewElement(pElement);
        }
    }
    else {
        DebugLog(L"Failed to retrieve UI element from point: " + std::to_wstring(hr));
    }
}

// Mouse hook procedure to detect mouse movements
LRESULT CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode >= 0 && wParam == WM_MOUSEMOVE) {
        POINT point;
        GetCursorPos(&point);
        ProcessCursorPosition(point);
        //auto now = std::chrono::steady_clock::now();
        //if (now - lastProcessed.load() > std::chrono::milliseconds(1)) {
        //    lastProcessed.store(now);
        //    ProcessCursorPosition(point);  // Directly call ProcessCursorPosition instead of async
        //}
    }
    return CallNextHookEx(hMouseHook, nCode, wParam, lParam);
}

// Reinitialize UI Automation and COM components
void ReinitializeAutomation() {
    DebugLog(L"Reinitializing UI Automation and COM components");

    if (pAutomation) {
        pAutomation.Release();
    }
    if (pPrevElement) {
        pPrevElement.Release();
    }
    {
        std::lock_guard<std::mutex> lock(pVoiceMtx);
        if (pVoice) {
            pVoice.Release();
        }
    }

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

    {
        std::lock_guard<std::mutex> lock(pVoiceMtx);
        pVoice->SetVolume(100);
        pVoice->SetRate(2);
    }

    DebugLog(L"Successfully reinitialized UI Automation and COM components");
}

// Schedule reinitialization task
void ScheduleReinitialization() {
    static boost::asio::steady_timer timer(io_context);

    timer.expires_after(std::chrono::minutes(10));
    timer.async_wait([&](const boost::system::error_code&) {
        io_context.post(ReinitializeAutomation);
        ScheduleReinitialization();
        });
}

// Shutdown function to clean up resources
void Shutdown() {
    DebugLog(L"Shutting down application.");

    UnhookWindowsHookEx(hMouseHook);

    std::lock_guard<std::mutex> lock(pVoiceMtx);
    pVoice.Release();

    pAutomation.Release();
    pPrevElement.Release();

    CoUninitialize();

    DebugLog(L"Application exited.");
}

// Mouse Input Thread Function
void MouseInputThread() {
    hMouseHook = SetWindowsHookEx(WH_MOUSE_LL, MouseProc, NULL, 0);
    if (!hMouseHook) {
        DebugLog(L"Failed to set mouse hook.");
        Shutdown();
        return;
    }

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    UnhookWindowsHookEx(hMouseHook);
}

// Updated Initialize Function
void Initialize() {
    DebugLog(L"Initialize function started.");
    try {
        SetConsoleBufferSize();
        _CrtSetDbgFlag(_CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF);
        if (_setmode(_fileno(stdout), _O_U16TEXT) == -1) throw std::runtime_error("Failed to set console mode");

        HRESULT hr = CoInitialize(NULL);
        if (FAILED(hr)) {
            DebugLog(L"Failed to initialize COM library: " + std::to_wstring(hr));
            throw std::runtime_error("Failed to initialize COM library");
        }

        pAutomation.Release();

        hr = CoCreateInstance(__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pAutomation));
        if (FAILED(hr)) {
            DebugLog(L"Failed to create UI Automation instance: " + std::to_wstring(hr));
            CoUninitialize();
            throw std::runtime_error("Failed to create UI Automation instance");
        }

        {
            std::lock_guard<std::mutex> lock(pVoiceMtx);

            pVoice.Release();

            hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void**)&pVoice);
            if (FAILED(hr)) {
                DebugLog(L"Failed to initialize SAPI: " + std::to_wstring(hr));
                pAutomation.Release();
                CoUninitialize();
                throw std::runtime_error("Failed to initialize SAPI");
            }

            pVoice->SetVolume(100);
            pVoice->SetRate(2);
        }

        // Start mouse input thread inside Initialize
        std::thread mouseThread(MouseInputThread);
        mouseThread.detach();

        ScheduleReinitialization();

        boost::asio::executor_work_guard<boost::asio::io_context::executor_type> work_guard(io_context.get_executor());

        std::thread keyboardHookThread(SetLowLevelKeyboardHook);
        DebugLog(L"Keyboard hook thread started.");
        keyboardHookThread.detach();
    }
    catch (const std::exception& e) {
        DebugLog(L"Exception in Initialize: " + Utf8ToWstring(e.what()));
        throw;
    }
}

// Updated Main Function
int main() {
    DebugLog(L"Main function started.");
    try {
        Initialize();
        DebugLog(L"Initialization successful.");

        boost::thread_group threads;
        for (std::size_t i = 0; i < std::thread::hardware_concurrency(); ++i) {
            threads.create_thread([&]() { io_context.run(); });
        }

        MSG msg;
        while (GetMessage(&msg, NULL, 0, 0)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }

        io_context.stop();
        threads.join_all();
        Shutdown();
    }
    catch (const std::exception& e) {
        DebugLog(L"Exception in main: " + Utf8ToWstring(e.what()));
        return 1;
    }

    return 0;
}

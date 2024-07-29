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

// ConcurrentQueue class definition to handle thread-safe queues
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
    mutable std::mutex mutex_;  // Mutex to protect the queue
    std::queue<T> queue_;       // Standard queue to hold elements
    std::condition_variable condition_;  // Condition variable for synchronization
};

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
    RECT rect;
};

// Global variables for UI Automation and speech synthesis
HHOOK hMouseHook; // Mouse hook handle
CComPtr<IUIAutomation> pAutomation = NULL; // UI Automation pointer
CComPtr<IUIAutomationElement> pPrevElement = NULL; // Previous UI element pointer
RECT prevRect = { 0, 0, 0, 0 }; // Previous rectangle
std::chrono::time_point<std::chrono::steady_clock> lastProcessedTime; // Last processed time
std::mutex mtx; // General mutex for synchronizing access

std::atomic<POINT> currentCursorPosition; // Current cursor position
std::atomic<bool> stopWorker(false); // Atomic boolean to signal worker thread to stop

CComPtr<ISpVoice> pVoice = NULL; // Speech voice pointer
std::mutex pVoiceMtx; // Mutex for pVoice

std::atomic<bool> cursorMoving(false); // Atomic boolean to track cursor movement
std::chrono::time_point<std::chrono::steady_clock> lastCursorMoveTime; // Last cursor move time

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
RECT currentRect; // Current rectangle being processed
std::mutex rectMtx; // Mutex for rectangle operations
std::condition_variable rectCv; // Condition variable for rectangle thread

// Struct to hold rectangle tasks (draw or clear)
struct RectangleTask {
    RECT rect; // Rectangle coordinates
    bool draw; // Boolean to indicate whether to draw or clear
};

ConcurrentQueue<RectangleTask> rectQueue; // Concurrent queue for rectangle tasks
std::atomic<bool> stopRectWorker(false); // Atomic boolean to signal rectangle worker thread to stop

// Worker threads for various tasks
std::thread worker;              // General worker thread
std::thread rectangleWorker;     // Thread for handling rectangle drawing and clearing
std::thread cursorChecker;       // Thread for checking cursor position

// Global variables for keyboard shortcuts
std::atomic<bool> capsLockOverride(false);
std::chrono::time_point<std::chrono::steady_clock> lastCapsLockPress;
std::mutex elementMutex;

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

void ProcessNewElement(CComPtr<IUIAutomationElement> pElement); // Function declaration

// Function to move to the parent element
void MoveToParentElement() {
    std::lock_guard<std::mutex> lock(elementMutex);
    if (!pPrevElement) return;

    CComPtr<IUIAutomationTreeWalker> pControlWalker;
    HRESULT hr = pAutomation->get_ControlViewWalker(&pControlWalker);
    if (SUCCEEDED(hr)) {
        CComPtr<IUIAutomationElement> pParent;
        hr = pControlWalker->GetParentElement(pPrevElement, &pParent);
        if (SUCCEEDED(hr) && pParent) {
            pPrevElement = pParent;
            DebugLog(L"Moved to parent element");
            ProcessNewElement(pPrevElement); // Process the new parent element
        }
    }
}

// Function to move to the first child element
void MoveToFirstChildElement() {
    std::lock_guard<std::mutex> lock(elementMutex);
    if (!pPrevElement) return;

    CComPtr<IUIAutomationTreeWalker> pControlWalker;
    HRESULT hr = pAutomation->get_ControlViewWalker(&pControlWalker);
    if (SUCCEEDED(hr)) {
        CComPtr<IUIAutomationElement> pChild;
        hr = pControlWalker->GetFirstChildElement(pPrevElement, &pChild);
        if (SUCCEEDED(hr) && pChild) {
            pPrevElement = pChild;
            DebugLog(L"Moved to first child element");
            ProcessNewElement(pPrevElement); // Process the new child element
        }
    }
}

// Add these functions to navigate to next sibling, previous sibling, and redo the current element
void MoveToNextSiblingElement() {
    std::lock_guard<std::mutex> lock(elementMutex);
    if (!pPrevElement) return;

    CComPtr<IUIAutomationTreeWalker> pControlWalker;
    HRESULT hr = pAutomation->get_ControlViewWalker(&pControlWalker);
    if (SUCCEEDED(hr)) {
        CComPtr<IUIAutomationElement> pNextSibling;
        hr = pControlWalker->GetNextSiblingElement(pPrevElement, &pNextSibling);
        if (SUCCEEDED(hr) && pNextSibling) {
            pPrevElement = pNextSibling;
            DebugLog(L"Moved to next sibling element");
            ProcessNewElement(pPrevElement); // Process the new sibling element
        }
    }
}

void MoveToPreviousSiblingElement() {
    std::lock_guard<std::mutex> lock(elementMutex);
    if (!pPrevElement) return;

    CComPtr<IUIAutomationTreeWalker> pControlWalker;
    HRESULT hr = pAutomation->get_ControlViewWalker(&pControlWalker);
    if (SUCCEEDED(hr)) {
        CComPtr<IUIAutomationElement> pPreviousSibling;
        hr = pControlWalker->GetPreviousSiblingElement(pPrevElement, &pPreviousSibling);
        if (SUCCEEDED(hr) && pPreviousSibling) {
            pPrevElement = pPreviousSibling;
            DebugLog(L"Moved to previous sibling element");
            ProcessNewElement(pPrevElement); // Process the new sibling element
        }
    }
}

void RedoCurrentElement() {
    std::lock_guard<std::mutex> lock(elementMutex);
    if (pPrevElement) {
        DebugLog(L"Redoing current element");
        ProcessNewElement(pPrevElement); // Reprocess the current element
    }
}

// Modify the LowLevelKeyboardProc function to handle new shortcuts
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

// Set the low-level keyboard hook
void SetLowLevelKeyboardHook() {
    HHOOK hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, LowLevelKeyboardProc, NULL, 0);
    if (!hKeyboardHook) {
        DebugLog(L"Failed to set keyboard hook.");
        return;
    }

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    UnhookWindowsHookEx(hKeyboardHook);
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
void QueueRectangle(const RECT& rect, bool draw) {
    if (rect.left == rect.right || rect.top == rect.bottom) {
        DebugLog(L"Ignoring invalid rectangle: (" + std::to_wstring(rect.left) + L", " + std::to_wstring(rect.top) + L", " + std::to_wstring(rect.right) + L", " + std::to_wstring(rect.bottom) + L")");
        return; // Ignore invalid rectangles
    }

    std::lock_guard<std::mutex> lock(rectMtx);  // Ensure exclusive access to the queue
    rectQueue.push({ rect, draw });
    rectCv.notify_one();

    std::wstring actionStr = draw ? L"Draw" : L"Clear";
    DebugLog(L"Queueing rectangle for " + actionStr + L": (" +
        std::to_wstring(rect.left) + L", " +
        std::to_wstring(rect.top) + L", " +
        std::to_wstring(rect.right) + L", " +
        std::to_wstring(rect.bottom) + L")");
}

// Process a rectangle task (draw or clear)
void ProcessRectangle(const RECT& rect, bool draw) {
    std::this_thread::sleep_for(std::chrono::milliseconds(125)); // Optional: wait before retrying
    try {
        DebugLog(L"Starting ProcessRectangle with action: " + std::to_wstring(draw) + L" for rectangle: (" +
            std::to_wstring(rect.left) + L", " +
            std::to_wstring(rect.top) + L", " +
            std::to_wstring(rect.right) + L", " +
            std::to_wstring(rect.bottom) + L")");

        HDC hdc = GetDC(NULL);
        if (hdc) {
            if (draw) {
                HPEN hPen = CreatePen(PS_SOLID, 2, RGB(255, 0, 0)); // Create a red pen for drawing
                HGDIOBJ hOldPen = SelectObject(hdc, hPen);
                HBRUSH hBrush = (HBRUSH)GetStockObject(HOLLOW_BRUSH);
                HGDIOBJ hOldBrush = SelectObject(hdc, hBrush);

                Rectangle(hdc, rect.left, rect.top, rect.right, rect.bottom); // Draw the rectangle

                SelectObject(hdc, hOldPen);
                SelectObject(hdc, hOldBrush);
                DeleteObject(hPen);

                rectangleDrawn.store(true);  // Update the state
                DebugLog(L"Rectangle drawn: (" + std::to_wstring(rect.left) + L", " +
                    std::to_wstring(rect.top) + L", " +
                    std::to_wstring(rect.right) + L", " +
                    std::to_wstring(rect.bottom) + L")");
            }
            else {
                InvalidateRect(NULL, &rect, TRUE); // Invalidate the rectangle area
                rectangleDrawn.store(false);  // Update the state

                DebugLog(L"Invalidated rectangle area: (" +
                    std::to_wstring(rect.left) + L", " +
                    std::to_wstring(rect.top) + L", " +
                    std::to_wstring(rect.right) + L", " +
                    std::to_wstring(rect.bottom) + L")");
            }
            ReleaseDC(NULL, hdc);
        }
        else {
            DebugLog(L"Failed to get device context.");
        }

        DebugLog(L"Ending ProcessRectangle with action: " + std::to_wstring(draw));
    }
    catch (const std::exception& e) {
        DebugLog(L"Exception in ProcessRectangle: " + Utf8ToWstring(e.what()));
    }
}

// Worker thread function to handle rectangle tasks
void RectangleWorker() {
    DebugLog(L"Starting RectangleWorker.");
    while (!stopRectWorker.load()) {
        try {
            RectangleTask task;
            if (rectQueue.try_pop(task)) {
                std::lock_guard<std::mutex> lock(rectMtx);  // Ensure exclusive access
                DebugLog(L"Processing rectangle: (" +
                    std::to_wstring(task.rect.left) + L", " +
                    std::to_wstring(task.rect.top) + L", " +
                    std::to_wstring(task.rect.right) + L", " +
                    std::to_wstring(task.rect.bottom) + L")");
                ProcessRectangle(task.rect, task.draw); // Process the rectangle task
            }
            else {
                std::unique_lock<std::mutex> lock(rectMtx);
                rectCv.wait(lock, [] { return !rectQueue.empty() || stopRectWorker.load(); });
            }
        }
        catch (const std::system_error& e) {
            DebugLog(L"System error in RectangleWorker: " + Utf8ToWstring(e.what()));
        }
        catch (const std::exception& e) {
            DebugLog(L"Exception in RectangleWorker: " + Utf8ToWstring(e.what()));
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

// Stop all current processes
void StopCurrentProcesses() {
    DebugLog(L"Entering StopCurrentProcesses.");

    try {
        newElementDetected.store(true); // Set the stop flag

        // Stop ongoing speech
        {
            std::scoped_lock lock(pVoiceMtx, speechMtx);
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
        speaking.store(false);
        speakingCv.notify_all();


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
        QueueRectangle(tempRect, false);

        // Clear the rect queue
        RectangleTask rectTask;
        while (rectQueue.try_pop(rectTask)) {}

        newElementDetected.store(false); // Reset the stop flag
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

            try {
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
            catch (const std::exception& e) {
                DebugLog(L"Exception while speaking text: " + Utf8ToWstring(e.what()));
                speaking.store(false);
                continue;
            }

            {
                std::lock_guard<std::mutex> rectLock(rectMtx);
                currentRect = textRect; // Set the current rectangle
            }
            QueueRectangle(currentRect, true); // Queue rectangle drawing

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

                std::this_thread::sleep_for(std::chrono::milliseconds(1)); // Sleep for a while before checking status again
            }

            QueueRectangle(currentRect, false); // Queue rectangle clearing

            DebugLog(L"Set speaking flag to false.");
            speaking.store(false); // Set speaking flag to false
            speakingCv.notify_all(); // Notify that speaking has finished
            DebugLog(L"Notified all waiting on speaking condition variable.");

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
    std::lock_guard<std::mutex> lock(mtx); // Lock mutex for thread safety
    if (pPrevElement == NULL && pElement == NULL) {
        return false;
    }
    else if (pPrevElement == NULL || pElement == NULL) {
        return true;
    }

    BOOL areSame;
    HRESULT hr = pAutomation->CompareElements(pPrevElement, pElement, &areSame); // Compare current and previous elements
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
    if (newElementDetected.load()) return; // Exit if processing is stopped

    try {
        CComPtr<IUIAutomationTextPattern> pTextPattern = NULL;
        HRESULT hr = pElement->GetCurrentPatternAs(UIA_TextPatternId, IID_PPV_ARGS(&pTextPattern)); // Get TextPattern from element
        if (SUCCEEDED(hr) && pTextPattern) {
            if (newElementDetected.load()) return; // Exit if processing is stopped
            std::this_thread::sleep_for(std::chrono::milliseconds(10)); // Delay to ensure the text is ready for extraction

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
    if (!pElement || newElementDetected.load()) return; // Exit if the element is null or processing is stopped

    struct ElementInfo {
        CComPtr<IUIAutomationElement> element;
        int depth;
    };

    ConcurrentQueue<ElementInfo> elementQueue;
    elementQueue.push({ pElement, 0 }); // Push the initial element with depth 0

    while (!elementQueue.empty() && elements.size() < maxElements) {
        if (newElementDetected.load()) return; // Exit if processing is stopped

        ElementInfo current;
        if (elementQueue.try_pop(current)) {
            if (newElementDetected.load()) return; // Exit if processing is stopped
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
                if (newElementDetected.load()) return; // Exit if processing is stopped

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
void ProcessNewElement(CComPtr<IUIAutomationElement> pElement) {
    StopCurrentProcesses();  // Ensure all current processes are stopped

    {
        std::lock_guard<std::mutex> lock(mtx);
        if (pPrevElement && rectangleDrawn) {
            QueueRectangle(prevRect, false); // Clear previous rectangle
        }

        HRESULT hr = pElement->get_CurrentBoundingRectangle(&prevRect); // Get bounding rectangle of new element
        if (SUCCEEDED(hr)) {
            QueueRectangle(prevRect, true); // Draw new rectangle
        }

        pPrevElement.Release();
        pPrevElement = pElement; // Update previous element
    }

    if (newElementDetected.load()) return; // Exit if a new element is detected

    std::vector<CComPtr<IUIAutomationElement>> elements;
    CollectElementsDFS(pElement, elements, MAX_DEPTH, MAX_CHILDREN, MAX_ELEMENTS); // Collect child elements

    ConcurrentQueue<TextRect> textRectQueue;
    size_t currentTextLength = 0;
    std::unordered_set<std::wstring> processedTexts;

    for (auto& element : elements) {
        if (newElementDetected.load()) return; // Exit if a new element is detected
        ReadElementText(element, textRectQueue, currentTextLength, processedTexts); // Read text from collected elements
    }

    if (newElementDetected.load()) return; // Exit if a new element is detected

    if (textRectQueue.empty()) {
        DebugLog(L"textRectQueue is empty after processing elements.");
    }
    else {
        DebugLog(L"textRectQueue populated with elements.");
    }

    std::wstring combinedText;

    {
        std::lock_guard<std::mutex> lock(speechMtx);

        TextRect temp;
        while (speechQueue.try_pop(temp)) {} // Clear speech queue

        while (textRectQueue.try_pop(temp)) {
            if (newElementDetected.load()) return; // Exit if a new element is detected
            if (!combinedText.empty()) {
                combinedText += L" ";
            }
            combinedText += temp.text; // Combine text from textRectQueue
            speechQueue.push(temp); // Transfer textRect to speechQueue
        }
    }

    if (!combinedText.empty()) {
        PrintText(combinedText); // Print combined text to console
    }

    speechCv.notify_all(); // Notify speech thread
    DebugLog(L"Notified speech thread.");
}


// Process cursor position and detect UI elements
void ProcessCursorPosition(POINT point) {
    DebugLog(L"Processing cursor position: (" + std::to_wstring(point.x) + L", " + std::to_wstring(point.y) + L")");

    CComPtr<IUIAutomationElement> pElement = NULL;
    HRESULT hr = pAutomation->ElementFromPoint(point, &pElement); // Get the UI element from the cursor position

    if (SUCCEEDED(hr) && pElement) {
        
        if (IsDifferentElement(pElement)) {
            DebugLog(L"New element detected.");
            StopCurrentProcesses(); // Stop all ongoing processes immediately
            ProcessNewElement(pElement); // Process the new UI element
        }
    }
    else {
        DebugLog(L"Failed to retrieve UI element from point: " + std::to_wstring(hr));
    }
}

// Worker thread function to process cursor movements
void WorkerThread() {
    POINT lastProcessedPoint = { -1, -1 }; // Initialize to an invalid point
    while (!stopWorker.load()) {
        POINT currentPoint = currentCursorPosition.load(); // Get the latest cursor position
        if (currentPoint.x != lastProcessedPoint.x || currentPoint.y != lastProcessedPoint.y) {
            ProcessCursorPosition(currentPoint); // Process the cursor position if it has changed
            lastProcessedPoint = currentPoint; // Update the last processed point
        }
        std::this_thread::sleep_for(std::chrono::milliseconds(10)); // Sleep to balance responsiveness and CPU usage
    }
}

// Mouse hook procedure to detect mouse movements
LRESULT CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode >= 0 && wParam == WM_MOUSEMOVE) {
        POINT point;
        GetCursorPos(&point);

        currentCursorPosition.store(point); // Update the current cursor position

        // Log the cursor movement
        DebugLog(L"Mouse moved to: (" + std::to_wstring(point.x) + L", " + std::to_wstring(point.y) + L")");
    }
    return CallNextHookEx(hMouseHook, nCode, wParam, lParam);
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
        return 2;
    }

    hr = CoCreateInstance(__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pAutomation)); // Create UI Automation instance
    if (FAILED(hr)) {
        DebugLog(L"Failed to create UI Automation instance: " + std::to_wstring(hr));
        CoUninitialize();
        return 3;
    }

    {
        std::lock_guard<std::mutex> lock(pVoiceMtx);
        hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void**)&pVoice); // Create SAPI voice instance
        if (FAILED(hr)) {
            DebugLog(L"Failed to initialize SAPI: " + std::to_wstring(hr));
            pAutomation.Release();
            CoUninitialize();
            return 4;
        }

        pVoice->SetVolume(100); // Set voice volume
        pVoice->SetRate(2); // Set voice rate
    }

    hMouseHook = SetWindowsHookEx(WH_MOUSE_LL, MouseProc, NULL, 0); // Set mouse hook
    if (!hMouseHook) {
        DebugLog(L"Failed to set mouse hook.");
        Shutdown(); // Shutdown if mouse hook fails
        return 5;
    }

    worker = std::thread(WorkerThread); // Start worker thread
    rectangleWorker = std::thread(RectangleWorker); // Start rectangle worker thread
    StartSpeechThread(); // Start speech thread

    // Start the keyboard hook in a separate thread
    std::thread keyboardHookThread(SetLowLevelKeyboardHook);
    keyboardHookThread.detach();

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

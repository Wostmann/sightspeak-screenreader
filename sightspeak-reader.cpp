#include <windows.h>
#include <WinSock2.h>
#include <Ws2tcpip.h>
#include <atlbase.h>
#include <UIAutomation.h>
#include <shellscalingapi.h>
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
#include <shared_mutex>
#include <future>
#include "external/BS_thread_pool.hpp"
#include "external/BS_thread_pool_utils.hpp"


// Utility function to convert UTF-8 string to wide string
// Used when converting text from multi-byte to wide-character encoding, common in Windows APIs
std::wstring Utf8ToWstring(const std::string& str) {
    if (str.empty()) return std::wstring();
    int size_needed = MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), NULL, 0);
    std::wstring wstrTo(size_needed, 0);
    MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), &wstrTo[0], size_needed);
    return wstrTo;
}

// Utility function to convert wide string to UTF-8 string
// Used when converting text from wide-character encoding back to multi-byte (UTF-8)
std::string WstringToUtf8(const std::wstring& wstr) {
    if (wstr.empty()) return std::string();
    int size_needed = WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), NULL, 0, NULL, NULL);
    std::string strTo(size_needed, 0);
    WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), &strTo[0], size_needed, NULL, NULL);
    return strTo;
}

// Structure to hold text and its associated rectangle
// Used to link textual content with its screen location for processing and drawing
struct TextRect {
    std::wstring text;
    RECT rect{ 0, 0, 0, 0 };
};

// Global variables for UI Automation and speech synthesis
HHOOK hMouseHook; // Hook for mouse input
CComPtr<IUIAutomation> pAutomation = NULL; // UI Automation instance
CComPtr<IUIAutomationElement> pPrevElement = NULL; // Previous UI element for comparison
RECT prevRect = { 0, 0, 0, 0 }; // Rectangle of the previous UI element
std::shared_mutex mtx; // Mutex for thread-safe access to shared resources
CComPtr<ISpVoice> pVoice = NULL; // SAPI voice instance for speech synthesis
std::mutex pVoiceMtx; // Mutex for thread-safe access to speech synthesis
std::atomic<bool> speaking(false); // Atomic flag indicating if speech is in progress
std::atomic<bool> rectangleDrawn(false); // Atomic flag indicating if a rectangle is drawn
const int MAX_DEPTH = 6; // Maximum depth for UI element traversal
COLORREF highlightColor = GetSysColor(COLOR_HIGHLIGHT); // Color used for highlighting elements
HPEN hPen = CreatePen(PS_SOLID, 2, highlightColor); // Pen used for drawing rectangles
HBRUSH hBrush = (HBRUSH)GetStockObject(HOLLOW_BRUSH); // Brush used for hollow rectangles
RECT currentRect; // Currently active rectangle
std::mutex rectMtx; // Mutex for thread-safe access to rectangle drawing
std::atomic<bool> capsLockOverride(false); // Atomic flag for CAPSLOCK override
std::atomic<bool> capsLockFirstPress{ false }; // Atomic flag for first CAPSLOCK press detection
std::chrono::time_point<std::chrono::steady_clock> lastCapsLockPress; // Timestamp of last CAPSLOCK press
std::shared_mutex elementMutex; // Shared mutex for UI element access
//boost::asio::io_context io_context; // Boost.Asio io_context for managing asynchronous tasks
BS::thread_pool pool(std::thread::hardware_concurrency());

std::atomic<int> taskVersion{ 0 };// Global atomic version counter to track task validity
std::unordered_set<std::wstring> processedTexts; // Set to keep track of processed texts

float scaleX = 1.0f;
float scaleY = 1.0f;

std::mutex cancelMtx; // Mutex for canceling tasks
std::promise<void> cancelPromise; // Promise to signal cancellation
std::shared_future<void> cancelFuture = cancelPromise.get_future().share(); // Future to share cancellation state

// Logging function to output debug messages to both debug console and log file
void DebugLog(const std::wstring& message) {
    std::wstringstream ws;
    ws << L"[DEBUG] " << message << std::endl;
    OutputDebugString(ws.str().c_str()); // Log to the debug output window
    std::wofstream logFile("debug.log", std::ios::app); // Log to a file for persistence
    logFile << ws.str();
}

// Function to toggle CAPSLOCK override state
// Allows CAPSLOCK to be used for navigation commands instead of its usual function
void ToggleCapsLockOverride() {
    capsLockOverride.store(!capsLockOverride.load()); // Toggle the override state
}

// Forward declaration of ProcessNewElement function
void ProcessNewElement(CComPtr<IUIAutomationElement> pElement);
void StopCurrentProcesses();

// Function to move to the parent element
// Used for navigating up the UI Automation tree
void MoveToParentElement() {
    std::shared_lock<std::shared_mutex> lock(elementMutex);  // Use shared_lock for read-only access
    if (!pPrevElement) return;

    CComPtr<IUIAutomationTreeWalker> pControlWalker;
    HRESULT hr = pAutomation->get_ControlViewWalker(&pControlWalker); // Get the tree walker for UI Automation
    if (SUCCEEDED(hr)) {
        CComPtr<IUIAutomationElement> pParent;
        hr = pControlWalker->GetParentElement(pPrevElement, &pParent); // Navigate to the parent element
        if (SUCCEEDED(hr) && pParent) {
            pPrevElement = pParent; // Update the current element to the parent
            ProcessNewElement(pPrevElement); // Process the new parent element
        }
    }
}

// Function to move to the first child element
// Used for navigating down the UI Automation tree to the first child
void MoveToFirstChildElement() {
    std::shared_lock<std::shared_mutex> lock(elementMutex);  // Use shared_lock for read-only access
    if (!pPrevElement) return;

    CComPtr<IUIAutomationTreeWalker> pControlWalker;
    HRESULT hr = pAutomation->get_ControlViewWalker(&pControlWalker); // Get the tree walker for UI Automation
    if (SUCCEEDED(hr)) {
        CComPtr<IUIAutomationElement> pChild;
        hr = pControlWalker->GetFirstChildElement(pPrevElement, &pChild); // Navigate to the first child element
        if (SUCCEEDED(hr) && pChild) {
            pPrevElement = pChild; // Update the current element to the first child
            ProcessNewElement(pPrevElement); // Process the new child element
        }
    }
}

// Function to move to the next sibling element
// Used for navigating across the UI Automation tree to the next sibling
void MoveToNextSiblingElement() {
    std::shared_lock<std::shared_mutex> lock(elementMutex);  // Use shared_lock for read-only access
    if (!pPrevElement) return;

    CComPtr<IUIAutomationTreeWalker> pControlWalker;
    HRESULT hr = pAutomation->get_ControlViewWalker(&pControlWalker); // Get the tree walker for UI Automation
    if (SUCCEEDED(hr)) {
        CComPtr<IUIAutomationElement> pNextSibling;
        hr = pControlWalker->GetNextSiblingElement(pPrevElement, &pNextSibling); // Navigate to the next sibling element
        if (SUCCEEDED(hr) && pNextSibling) {
            pPrevElement = pNextSibling; // Update the current element to the next sibling
            ProcessNewElement(pPrevElement); // Process the new sibling element
        }
    }
}

// Function to move to the previous sibling element
// Used for navigating across the UI Automation tree to the previous sibling
void MoveToPreviousSiblingElement() {
    std::shared_lock<std::shared_mutex> lock(elementMutex);  // Use shared_lock for read-only access
    if (!pPrevElement) return;

    CComPtr<IUIAutomationTreeWalker> pControlWalker;
    HRESULT hr = pAutomation->get_ControlViewWalker(&pControlWalker); // Get the tree walker for UI Automation
    if (SUCCEEDED(hr)) {
        CComPtr<IUIAutomationElement> pPreviousSibling;
        hr = pControlWalker->GetPreviousSiblingElement(pPrevElement, &pPreviousSibling); // Navigate to the previous sibling element
        if (SUCCEEDED(hr) && pPreviousSibling) {
            pPrevElement = pPreviousSibling; // Update the current element to the previous sibling
            ProcessNewElement(pPrevElement); // Process the new sibling element
        }
    }
}

// Function to redo the current element
// Reprocesses the current UI element to reapply any relevant processing or drawing
void RedoCurrentElement() {
    std::shared_lock<std::shared_mutex> lock(elementMutex);  // Use shared_lock for read-only access
    if (pPrevElement) {
        ProcessNewElement(pPrevElement); // Reprocess the current element
    }
}

// Low-level keyboard procedure to handle keyboard shortcuts
// Hooks into keyboard input to detect and handle specific key combinations for navigation
LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode == HC_ACTION) {
        KBDLLHOOKSTRUCT* pKeyBoard = (KBDLLHOOKSTRUCT*)lParam; // Cast the lParam to KBDLLHOOKSTRUCT to access keyboard event data
        if (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN) {

            // Detecting CTRL key press
            if (pKeyBoard->vkCode == VK_LCONTROL || pKeyBoard->vkCode == VK_RCONTROL) { // Check for left or right CTRL
                StopCurrentProcesses(); // Stop all processes
            }

            if (pKeyBoard->vkCode == VK_CAPITAL) {
                auto now = std::chrono::steady_clock::now();
                if (std::chrono::duration_cast<std::chrono::milliseconds>(now - lastCapsLockPress).count() < 500 && capsLockFirstPress.load()) {
                    ToggleCapsLockOverride(); // Toggle CAPSLOCK override if double pressed within 500ms
                    capsLockFirstPress = false;  // Reset the flag after the second press
                }
                else {
                    capsLockFirstPress = true;  // Set the flag for the first press
                }
                lastCapsLockPress = now;
            }
            else if (capsLockFirstPress.load()) {
                // If any other key is pressed after the first CAPSLOCK press, reset the flag
                capsLockFirstPress = false;
            }

            if (capsLockOverride.load()) { // Check if CAPSLOCK override is enabled
                switch (pKeyBoard->vkCode) {
                case 'W':
                    MoveToParentElement(); // Navigate to parent element on CAPSLOCK+W
                    break;
                case 'S':
                    MoveToFirstChildElement(); // Navigate to first child element on CAPSLOCK+S
                    break;
                case 'D':
                    MoveToNextSiblingElement(); // Navigate to next sibling element on CAPSLOCK+D
                    break;
                case 'A':
                    MoveToPreviousSiblingElement(); // Navigate to previous sibling element on CAPSLOCK+A
                    break;
                case 'E':
                    RedoCurrentElement(); // Redo the current element on CAPSLOCK+E
                    break;
                }
            }
        }
    }
    return CallNextHookEx(NULL, nCode, wParam, lParam); // Pass the event to the next hook in the chain
}


// Function to set the low-level keyboard hook
// Hooks into the keyboard input stream to intercept keypresses for custom handling
void SetLowLevelKeyboardHook() {

    HHOOK hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, LowLevelKeyboardProc, NULL, 0);
    if (!hKeyboardHook) {
        DebugLog(L"Failed to set keyboard hook. Error: " + std::to_wstring(GetLastError()));
        return;
    }

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessage(&msg); // Process Windows messages
    }

    UnhookWindowsHookEx(hKeyboardHook); // Unhook the keyboard hook when the message loop exits
}

// Function to set the console buffer size for better visibility of console output
// Expands the console buffer to accommodate large amounts of output data
void SetConsoleBufferSize() {
    CONSOLE_SCREEN_BUFFER_INFOEX csbiex = { 0 };
    csbiex.cbSize = sizeof(csbiex); // Initialize the structure size
    HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
    if (GetConsoleScreenBufferInfoEx(hConsole, &csbiex)) {
        csbiex.dwSize.X = 8000; // Set the buffer width
        csbiex.dwSize.Y = 2000; // Set the buffer height
        SetConsoleScreenBufferInfoEx(hConsole, &csbiex); // Apply the new buffer size
    }
}

// Function to process a rectangle on the screen
// Draws or clears a rectangle around the specified area
void ProcessRectangle(const RECT& rect, bool draw, std::shared_future<void> cancelFuture) {
    if (cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) { return; } // Exit if cancellation is requested

    // Get the DPI scaling factor
    RECT scaledRect = {
        static_cast<LONG>(rect.left * scaleX),
        static_cast<LONG>(rect.top * scaleY),
        static_cast<LONG>(rect.right * scaleX),
        static_cast<LONG>(rect.bottom * scaleY)
    };

    try {
        HDC hdc = GetDC(NULL); // Get the device context for drawing on the screen
        if (hdc) {
            if (draw) {
                // Delay handling is managed outside the critical section to avoid locking overhead
                HGDIOBJ hOldPen = SelectObject(hdc, hPen); // Select the pen for drawing
                HGDIOBJ hOldBrush = SelectObject(hdc, hBrush); // Select the brush for drawing

                if (cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) {
                    ReleaseDC(NULL, hdc); // Release the device context before returning
                    return;
                } // Exit if cancellation is requested

                std::this_thread::sleep_for(std::chrono::milliseconds(30)); // Sleep to simulate drawing delay

                if (cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) {
                    ReleaseDC(NULL, hdc); // Release the device context before returning
                    return;
                } // Exit if cancellation is requested

                std::this_thread::sleep_for(std::chrono::milliseconds(30)); // Sleep to simulate drawing delay

                if (cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) {
                    ReleaseDC(NULL, hdc); // Release the device context before returning
                    return;
                } // Exit if cancellation is requested

                Rectangle(hdc, scaledRect.left, scaledRect.top, scaledRect.right, scaledRect.bottom); // Draw the rectangle

                SelectObject(hdc, hOldPen); // Restore the previous pen
                SelectObject(hdc, hOldBrush); // Restore the previous brush
                rectangleDrawn.store(true); // Mark the rectangle as drawn
                {
                    // Lock the rectangle resource to ensure thread-safe access
                    std::lock_guard<std::mutex> rectLock(rectMtx);
                    currentRect = rect; // Update the current rectangle with the scaled rectangle
                }
            }
            else {
                InvalidateRect(NULL, &scaledRect, TRUE); // Clear the rectangle
                rectangleDrawn.store(false); // Mark the rectangle as not drawn
            }
            ReleaseDC(NULL, hdc); // Release the device context
        }
        else {
            DebugLog(L"Failed to get device context."); // Log failure if device context is not available
        }
    }
    catch (const std::exception& e) {
        DebugLog(L"Exception in ProcessRectangle: " + Utf8ToWstring(e.what())); // Log any exceptions that occur
    }
}


// Function to print text to the console and log it
// Outputs text to the console and also logs it for debugging
void PrintText(const std::wstring& text) {
    std::wcout << text << std::endl; // Output the text to the console
    std::wcout.flush(); // Flush the console output
}

// Task to speak text and manage rectangle
// Asynchronously processes text for speech and manages the associated rectangle
void SpeakTextTask(const std::wstring& textToSpeak, std::shared_future<void> cancelFuture) {
    // Check if the task should be canceled before starting
    if (cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) { return; } // Exit if cancellation is requested

    try {
        // If the text is empty or cancellation is requested, exit early
        if (cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) { return; } // Exit if cancellation is requested

        PrintText(textToSpeak); // Output the text to the console and log it

        speaking.store(true); // Set the speaking flag to true, indicating speech is in progress

        {
            // Lock the speech synthesis resource to ensure thread-safe access
            std::lock_guard<std::mutex> lock(pVoiceMtx);
            if (pVoice) { // Check if the speech synthesis object is valid
                // Check for cancellation again before starting speech
                if (cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) { return; } // Exit if cancellation is requested

                // Purge any previous speech to start fresh
                HRESULT hr = pVoice->Speak(nullptr, SPF_PURGEBEFORESPEAK, nullptr);
                if (FAILED(hr)) {
                    DebugLog(L"Failed to purge speech: " + std::to_wstring(hr)); // Log failure to purge
                    speaking.store(false); // Reset the speaking flag if purging failed
                    return;
                }

                // Start speaking the text asynchronously
                hr = pVoice->Speak(textToSpeak.c_str(), SPF_ASYNC, nullptr);
                if (FAILED(hr)) {
                    DebugLog(L"Failed to speak text: " + textToSpeak + L" Error: " + std::to_wstring(hr)); // Log failure to speak
                    speaking.store(false); // Reset the speaking flag if speaking failed
                    return;
                }
            }
        }

        SPVOICESTATUS status; // Structure to hold the status of the speech synthesis
        while (true) {
            // Check for cancellation every 10 milliseconds
            if (cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) { return; } // Exit if cancellation is requested

            // Lock the speech synthesis resource to check its status
            std::lock_guard<std::mutex> lock(pVoiceMtx);
            if (pVoice) { // Check if the speech synthesis object is valid
                HRESULT hr = pVoice->GetStatus(&status, nullptr); // Get the current status of the speech synthesis
                if (FAILED(hr)) {
                    DebugLog(L"Failed to get status: " + std::to_wstring(hr)); // Log failure to get status
                    break; // Exit the loop if getting the status failed
                }

                // Check if the speech synthesis has completed
                if (status.dwRunningState == SPRS_DONE) {
                    break; // Exit the loop if the speech is done
                }
            }
        }

        speaking.store(false); // Reset the speaking flag to indicate speech is complete
    }
    catch (const std::exception& e) {
        // Log any exceptions that occur during the task
        DebugLog(L"Exception in SpeakTextTask: " + Utf8ToWstring(e.what()));
    }
}


// Class to manage the queue for processing TextRect objects
// Handles the queuing and processing of text and associated rectangles asynchronously
class ProcessTextRectQueue {
public:
    // Enqueue a TextRect object for processing
    // This function adds a TextRect to the queue and triggers processing if not already speaking
    static void Enqueue(TextRect textRect, std::shared_future<void> cancelFuture) {
        if (cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) { return; } // Exit if cancellation is requested
        {
            std::lock_guard<std::mutex> lock(queueMutex);
            textRectQueue.push(textRect); // Add the TextRect to the queue
        }
        if (!speaking.load()) {
            pool.detach_task(std::bind(&ProcessTextRectQueue::DequeueAndProcess, cancelFuture)); // Start processing if not already speaking
        }
    }

    // Dequeue and process the next TextRect in the queue
// Handles the drawing and speaking of the text and rectangle
    static void DequeueAndProcess(std::shared_future<void> cancelFuture) {
        if (textRectQueue.empty() || cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) {
            speaking.store(false); // Reset the speaking flag if queue is empty or canceled
            return;
        }

        speaking.store(true); // Set the speaking flag to true
        TextRect textRect;
        {
            std::lock_guard<std::mutex> lock(queueMutex);
            textRect = textRectQueue.front(); // Get the next TextRect from the queue
            textRectQueue.pop(); // Remove it from the queue
        }

        // Process the rectangle drawing first
        auto rectFuture = pool.submit_task([&]() {ProcessRectangle(textRect.rect, true, cancelFuture); });
        // Handle the text and rectangle in a non-blocking way
        auto future = pool.submit_task([&]() {
            SpeakTextTask(textRect.text, cancelFuture); // Speak the text
            });

        future.get(); // Wait for the speaking task to complete
        rectFuture.get();
        // Clear the rectangle after speaking is done
        ProcessRectangle(textRect.rect, false, cancelFuture); // Clear the rectangle

        // Reset the speaking flag after task completion
        pool.detach_task(std::bind(&ProcessTextRectQueue::DequeueAndProcess, cancelFuture)); // Trigger the next item in the queue
    }

    // Clear the TextRect queue
    // Empties the queue and ensures the speaking flag is reset
    static void ClearQueue() {
        std::lock_guard<std::mutex> lock(queueMutex);
        std::queue<TextRect> empty;
        std::swap(textRectQueue, empty); // Swap with an empty queue to clear it
        speaking.store(false);  // Ensure that speaking flag is reset
    }

private:
    static std::queue<TextRect> textRectQueue; // Queue to hold TextRect objects for processing
    static std::mutex queueMutex; // Mutex for thread-safe access to the queue
    static std::atomic<bool> speaking; // Atomic flag indicating if speech is in progress
};

std::queue<TextRect> ProcessTextRectQueue::textRectQueue; // Initialize the static textRectQueue
std::mutex ProcessTextRectQueue::queueMutex; // Initialize the static queue mutex
std::atomic<bool> ProcessTextRectQueue::speaking = false; // Initialize the static speaking flag

// Function to read text and rectangle from a UI element
// Extracts text and bounding rectangles from a UI element for processing
void ReadElementText(CComPtr<IUIAutomationElement> pElement, std::shared_future<void> cancelFuture) {
    if (cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) { return; } // Exit if cancellation is requested

    try {
        // Process the text content and bounding rectangle
        CComPtr<IUIAutomationTextPattern> pTextPattern = NULL;
        HRESULT hr = pElement->GetCurrentPatternAs(UIA_TextPatternId, IID_PPV_ARGS(&pTextPattern)); // Get the text pattern from the element
        if (SUCCEEDED(hr) && pTextPattern) {

            CComPtr<IUIAutomationTextRange> pTextRange = NULL;
            hr = pTextPattern->get_DocumentRange(&pTextRange); // Get the text range from the text pattern
            if (SUCCEEDED(hr) && pTextRange) {
                CComBSTR text;
                hr = pTextRange->GetText(-1, &text); // Get the text within the text range
                if (SUCCEEDED(hr) && text != NULL) {
                    std::wstring textStr(static_cast<wchar_t*>(text)); // Convert the BSTR text to std::wstring
                    if (!textStr.empty() && processedTexts.find(textStr) == processedTexts.end()) {
                        RECT rect = {};
                        hr = pElement->get_CurrentBoundingRectangle(&rect); // Get the bounding rectangle of the UI element
                        if (SUCCEEDED(hr)) {
                            ProcessTextRectQueue::Enqueue({ textStr, rect }, cancelFuture); // Enqueue the text and rectangle for processing
                            processedTexts.insert(textStr);  // Add the text to the set after enqueueing to avoid reprocessing
                        }
                    }
                }
            }
        }

        // Process the name and bounding rectangle
        CComBSTR name;
        hr = pElement->get_CurrentName(&name); // Get the name property of the UI element
        if (SUCCEEDED(hr) && name != NULL) {
            std::wstring nameStr(static_cast<wchar_t*>(name)); // Convert the BSTR name to std::wstring
            if (!nameStr.empty() && processedTexts.find(nameStr) == processedTexts.end()) {
                RECT rect = {};
                hr = pElement->get_CurrentBoundingRectangle(&rect); // Get the bounding rectangle of the UI element
                if (SUCCEEDED(hr)) {
                    ProcessTextRectQueue::Enqueue({ nameStr, rect }, cancelFuture); // Enqueue the name and rectangle for processing
                    processedTexts.insert(nameStr);  // Add the name to the set after enqueueing to avoid reprocessing
                }
            }
        }
    }
    catch (const std::exception& e) {
        DebugLog(L"Exception in ReadElementText: " + Utf8ToWstring(e.what())); // Log any exceptions that occur
    }
}


// Collect UI elements using breadth-first search
// Traverses the UI Automation tree to gather elements and process their text and rectangles
void CollectElementsBFS(CComPtr<IUIAutomationElement> pElement, std::shared_future<void> cancelFuture) {
    if (!pElement || cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) return;
    processedTexts.clear(); // Clear the set of processed texts to start fresh
    struct ElementInfo {
        CComPtr<IUIAutomationElement> element; // The UI element to process
        int depth{ 0 }; // The depth of the element in the UI tree
    };

    std::queue<ElementInfo> elementQueue;
    elementQueue.push({ pElement, 0 }); // Start with the root element

    while (!elementQueue.empty()) {
        if (cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) return;

        ElementInfo current = std::move(elementQueue.front()); // Get the next element in the queue
        elementQueue.pop();
        if (current.depth >= MAX_DEPTH) continue; // Skip elements that are too deep

        if (cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) return;
        ReadElementText(current.element, cancelFuture); // Process the text and rectangle of the element

        CComPtr<IUIAutomationTreeWalker> pControlWalker;
        HRESULT hr = pAutomation->get_ControlViewWalker(&pControlWalker); // Get the tree walker for UI Automation
        if (FAILED(hr)) {
            DebugLog(L"Failed to get ControlViewWalker: " + std::to_wstring(hr)); // Log failure to get tree walker
            continue;
        }

        CComPtr<IUIAutomationElement> pChild;
        hr = pControlWalker->GetFirstChildElement(current.element, &pChild); // Get the first child element
        if (FAILED(hr)) {
            DebugLog(L"Failed to get first child element: " + std::to_wstring(hr)); // Log failure to get child element
            continue;
        }

        while (SUCCEEDED(hr) && pChild) { // Check if there are children
            if (cancelFuture.wait_for(std::chrono::seconds(0)) == std::future_status::ready) return;

            elementQueue.push({ pChild, current.depth + 1 }); // Add the child element to the queue

            CComPtr<IUIAutomationElement> pNextSibling;
            hr = pControlWalker->GetNextSiblingElement(pChild, &pNextSibling); // Get the next sibling element
            if (FAILED(hr)) {
                DebugLog(L"Failed to get next sibling element: " + std::to_wstring(hr)); // Log failure to get sibling element
                break;
            }

            pChild = pNextSibling; // Move to the next sibling
        }
    }
}

// Function to stop current processes asynchronously
// Cancels all active tasks and clears the processing queue
void StopCurrentProcesses() {

    try {
        // Signal cancellation of current tasks
        {
            std::lock_guard<std::mutex> cancelLock(cancelMtx);
            cancelPromise.set_value(); // Set the promise to cancel current tasks
            cancelPromise = std::promise<void>();  // Reset for new tasks
            cancelFuture = cancelPromise.get_future().share(); // Create a new shared future for cancellation
        }

        // Stop ongoing processes
        ProcessTextRectQueue::ClearQueue(); // Clear the processing queue

        if (rectangleDrawn.load()) {  // Clear the rectangle if one is drawn
            std::lock_guard<std::mutex> rectLock(rectMtx);
            pool.detach_task([=]() {ProcessRectangle(currentRect, false, cancelFuture); }); // Clear the rectangle asynchronously
        }

        {   //Clear speech
            std::lock_guard<std::mutex> voiceLock(pVoiceMtx);
            if (pVoice) {
                HRESULT hr = pVoice->Speak(nullptr, SPF_PURGEBEFORESPEAK, nullptr); // Purge any ongoing speech
                if (FAILED(hr)) {
                    DebugLog(L"Failed to stop speech: " + std::to_wstring(hr)); // Log failure to stop speech
                }
            }
        }

    }
    catch (const std::system_error& e) {
        DebugLog(L"Exception in StopCurrentProcessesAsync: " + Utf8ToWstring(e.what())); // Log any system exceptions
    }
    catch (const std::exception& e) {
        DebugLog(L"Exception in StopCurrentProcessesAsync: " + Utf8ToWstring(e.what())); // Log any general exceptions
    }
}



// Process new UI element
// Handles the detection and processing of new UI elements under the cursor
void ProcessNewElement(CComPtr<IUIAutomationElement> pElement) {
    // Increment the task version to invalidate all previous tasks
    int currentVersion = ++taskVersion;
    
    pool.detach_task([pElement, currentVersion]() {
        // If the current task version is outdated, skip this task
        if (currentVersion != taskVersion.load()) {
            return;
        }

        StopCurrentProcesses();  // Stop all current tasks

        // If the task is still valid, proceed with BFS to collect UI elements
        if (currentVersion == taskVersion.load()) {
            CollectElementsBFS(pElement, cancelFuture);
        }
        });
}


// Check if the UI element is different from the previous one
// Compares the current UI element with the previous one to detect changes
bool IsDifferentElement(CComPtr<IUIAutomationElement> pElement) {
    std::shared_lock<std::shared_mutex> lock(mtx);  // Use shared_lock for read-only access
    if (pPrevElement == NULL && pElement == NULL) {
        return false; // No elements to compare
    }
    else if (pPrevElement == NULL || pElement == NULL) {
        return true; // One of the elements is NULL, so they are different
    }

    BOOL areSame;
    HRESULT hr = pAutomation->CompareElements(pPrevElement, pElement, &areSame); // Compare the elements using UI Automation
    return SUCCEEDED(hr) && !areSame; // Return true if the elements are different
}

// Process cursor position and detect UI elements
// Retrieves the UI element under the cursor and triggers processing if it has changed
void ProcessCursorPosition(POINT point) {
    CComPtr<IUIAutomationElement> pElement = NULL;
    HRESULT hr = pAutomation->ElementFromPoint(point, &pElement); // Get the UI element under the cursor

    if (SUCCEEDED(hr) && pElement) {
        if (IsDifferentElement(pElement)) { // Check if the element is different from the previous one
            pPrevElement.Release(); // Release the previous element
            pPrevElement = pElement; // Update the previous element to the current one
            ProcessNewElement(pElement); // Process the new element
        }
    }
    else {
        DebugLog(L"Failed to retrieve UI element from point: " + std::to_wstring(hr)); // Log failure to retrieve the element
    }
}

// Mouse hook procedure to detect mouse movements
// Monitors mouse movements to detect when the cursor is over a new UI element
LRESULT CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode >= 0 && wParam == WM_MOUSEMOVE) {
        POINT point;
        GetCursorPos(&point); // Get the current cursor position
        pool.detach_task([point]() {ProcessCursorPosition(point);}); // Process the cursor position to detect UI elements
    }
    return CallNextHookEx(hMouseHook, nCode, wParam, lParam); // Pass the event to the next hook in the chain
}

// Reinitialize UI Automation and COM components
// Resets and reinitializes UI Automation and speech synthesis components periodically
void ReinitializeAutomation() {

    if (pAutomation) {
        pAutomation.Release(); // Release the existing UI Automation instance
    }
    if (pPrevElement) {
        pPrevElement.Release(); // Release the previous UI element
    }
    {
        std::lock_guard<std::mutex> lock(pVoiceMtx);
        if (pVoice) {
            pVoice.Release(); // Release the existing speech synthesis instance
        }
    }

    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) {
        DebugLog(L"Failed to reinitialize COM library: " + std::to_wstring(hr)); // Log failure to reinitialize COM
        return;
    }

    hr = CoCreateInstance(__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pAutomation)); // Create a new UI Automation instance
    if (FAILED(hr)) {
        DebugLog(L"Failed to reinitialize UI Automation: " + std::to_wstring(hr)); // Log failure to reinitialize UI Automation
        CoUninitialize();
        return;
    }

    hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void**)&pVoice); // Create a new speech synthesis instance
    if (FAILED(hr)) {
        DebugLog(L"Failed to reinitialize SAPI: " + std::to_wstring(hr)); // Log failure to reinitialize SAPI
        pAutomation.Release();
        CoUninitialize();
        return;
    }

    {
        std::lock_guard<std::mutex> lock(pVoiceMtx);
        pVoice->SetVolume(100); // Set the volume of the speech synthesis
        pVoice->SetRate(2); // Set the rate of the speech synthesis
    }
}

// Schedule reinitialization task
// Sets up a periodic task to reinitialize UI Automation and speech synthesis components
void ScheduleReinitialization() {
    // Submit the reinitialization task with a delay
    pool.detach_task([]() {
        while (true) {
            std::this_thread::sleep_for(std::chrono::minutes(10)); // Wait for 10 minutes
            ReinitializeAutomation(); // Reinitialize the components
        }
        }); // Detach the task so it runs independently
}


// Shutdown function to clean up resources
// Handles the clean-up of hooks, COM objects, and other resources before exiting the application
void Shutdown() {

    UnhookWindowsHookEx(hMouseHook); // Unhook the mouse hook

    std::lock_guard<std::mutex> lock(pVoiceMtx);
    pVoice.Release(); // Release the speech synthesis instance

    pAutomation.Release(); // Release the UI Automation instance
    pPrevElement.Release(); // Release the previous UI element
    DeleteObject(hPen); // Delete the pen used for drawing rectangles

    CoUninitialize(); // Uninitialize COM

}

// Mouse Input Thread Function
// Runs in a separate thread to handle mouse input and detect UI elements under the cursor
void MouseInputThread() {
    hMouseHook = SetWindowsHookEx(WH_MOUSE_LL, MouseProc, NULL, 0);
    if (!hMouseHook) {
        DebugLog(L"Failed to set mouse hook."); // Log failure to set the mouse hook
        Shutdown(); // Shutdown the application if the mouse hook cannot be set
        return;
    }

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg); // Process Windows messages
    }

    UnhookWindowsHookEx(hMouseHook); // Unhook the mouse hook when the message loop exits
}

// Updated Initialize Function
// Sets up the initial state, including console buffer, COM initialization, and hook setup
void Initialize() {
    try {
        SetConsoleBufferSize(); // Set the console buffer size for better output visibility
        _CrtSetDbgFlag(_CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF); // Enable memory leak detection
        if (_setmode(_fileno(stdout), _O_U16TEXT) == -1) throw std::runtime_error("Failed to set console mode"); // Set the console mode for wide character output
        HRESULT hrDPI = SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE);
        HDC hdc = GetDC(NULL);
        int dpiX = GetDeviceCaps(hdc, LOGPIXELSX);
        int dpiY = GetDeviceCaps(hdc, LOGPIXELSY);
        float scaleX = dpiX / 96.0f;
        float scaleY = dpiY / 96.0f;
        ReleaseDC(NULL, hdc);

        if (FAILED(hrDPI)) {
            DebugLog(L"Failed to set DPI awareness: " + std::to_wstring(hrDPI));
        }
        HRESULT hr = CoInitialize(NULL); // Initialize COM
        if (FAILED(hr)) {
            DebugLog(L"Failed to initialize COM library: " + std::to_wstring(hr)); // Log failure to initialize COM
            throw std::runtime_error("Failed to initialize COM library");
        }

        pAutomation.Release(); // Release any existing UI Automation instance

        hr = CoCreateInstance(__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pAutomation)); // Create a new UI Automation instance
        if (FAILED(hr)) {
            DebugLog(L"Failed to create UI Automation instance: " + std::to_wstring(hr)); // Log failure to create UI Automation instance
            CoUninitialize();
            throw std::runtime_error("Failed to create UI Automation instance");
        }

        {
            std::lock_guard<std::mutex> lock(pVoiceMtx);

            pVoice.Release(); // Release any existing speech synthesis instance

            hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void**)&pVoice); // Create a new speech synthesis instance
            if (FAILED(hr)) {
                DebugLog(L"Failed to initialize SAPI: " + std::to_wstring(hr)); // Log failure to initialize SAPI
                pAutomation.Release();
                CoUninitialize();
                throw std::runtime_error("Failed to initialize SAPI");
            }

            pVoice->SetVolume(100); // Set the volume of the speech synthesis
            pVoice->SetRate(2); // Set the rate of the speech synthesis
        }

        // Start mouse input thread inside Initialize
        std::thread mouseThread(MouseInputThread); // Start the mouse input thread
        mouseThread.detach(); // Detach the thread to allow it to run independently

        ScheduleReinitialization(); // Schedule periodic reinitialization of components

        //boost::asio::executor_work_guard<boost::asio::io_context::executor_type> work_guard(io_context.get_executor()); // Prevent the io_context from running out of work

        std::thread keyboardHookThread(SetLowLevelKeyboardHook); // Start the keyboard hook thread
        keyboardHookThread.detach(); // Detach the thread to allow it to run independently
    }
    catch (const std::exception& e) {
        DebugLog(L"Exception in Initialize: " + Utf8ToWstring(e.what())); // Log any exceptions during initialization
        throw;
    }
}

// Entry point of the application, handles initialization, message loop, and shutdown
int main() {

    try {
        Initialize(); // Initialize the application

        // Schedule any initial tasks needed using the thread pool
        // For example, if you need to start periodic reinitialization:
        ScheduleReinitialization();

        MSG msg;
        while (GetMessage(&msg, NULL, 0, 0)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg); // Process Windows messages
        }

        pool.wait(); // Wait for all tasks to complete before shutting down
        Shutdown(); // Shutdown the application
    }
    catch (const std::exception& e) {
        DebugLog(L"Exception in main: " + Utf8ToWstring(e.what())); // Log any exceptions in the main function
        return 1;
    }

    return 0; // Exit the application with a success code
}
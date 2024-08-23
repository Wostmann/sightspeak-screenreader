# SightSpeak Reader

## Overview

**SightSpeak Reader** is a custom screen reader designed to assist visually impaired users by utilizing cursor movement detection to read out and outline the text under the cursor as it hovers over elements on the screen. This project is currently in beta release and is under active development. The screen reader is built using C++ and integrates various libraries such as WinAPI, Microsoft UI Automation, Windows Speech API, and BS_thread_pool for efficient program management and to ensure a seamless reading experience across all Microsoft applications and environments.

## Key Features

- **Text-to-Speech (TTS) Output:** The screen reader reads aloud the text content of UI elements when hovered over by the cursor.
- **Text Highlighting:** Visually outlines the text being read, synchronizing the spoken text with the highlighted text.
- **Top-Down Reading:** The screen reader is able to extract large amounts of text through a BFS UI tree traversal. It then is able to sequentially output these collected elements from the top down.
- **Keyboard Commands:** Built in keyboard shortcuts exist in the program to allow for manual traversal of the UI tree as well as full control of stopping and pausing of speech.
- **Multi-Threaded Processing:** Efficiently handles multiple tasks simultaneously to ensure smooth operation and immediate response to user input.

## Beta Release Notice
This project is currently in its beta release phase. While the core functionality is in place, there are further improvements and issues that are still being addressed. We encourage users to provide feedback to help improve the software.
## Installation

### Prerequisites

- Windows 10/11
- Visual Studio 2022 or alternative C++ compiler

### Building the Project

1. Option 1
   
Run the executable located in the Releases section of the GitHub repository.
The program will immediately begin processing UI elements under the cursor, highlighting text and reading it aloud.

2. Option 2
   
If you want to customize or have access to the code, clone the repository and build the project in release mode.

### Navigating UI Elements

- Use the predefined key bindings to navigate through different UI elements on the screen:
- CAPSLOCK + CAPSLOCK: Double press CAPSLOCK to toggle the override on the CAPSLOCK button to start using the shortcuts.
- CAPSLOCK + W: Move up to the parent UI element.
- CAPSLOCK + S: Move down to the first child UI element.
- CAPSLOCK + D: Move to the next sibling UI element.
- CAPSLOCK + A: Move to the previous sibling UI element.
- CAPSLOCK + E: Re-read the current UI element.
- CAPSLOCK + Q: Quit the program.
- CTRL: Pause the program.
These commands allow for efficient navigation through UI elements and control over the reading process.

## Future Improvements

- Windows Magnifier Interaction: In the future, the program aims to integrate with the Windows Magnifier API so that rectangle drawing and resizing will be done properly.
- Tesseract OCR: We plan to add a feature where, upon command or detection of a PDF/image, the cursor position will be sent to Tesseract's OCR processing to extract text from places not using accessibility APIs
- Outline Refining: The problem where rectangles will be drawn and then prematurely deleted on some elements will be debugged.
- Application Detection: While using SightSpeak, we hope to be able to tell what applications are being run such as if another screen reader is being run, or which web browser is currently active to avoid errors and improve efficiency. 

## Contributions

We welcome contributions to the project. Please fork the repository and submit a pull request with your changes. Be sure to include detailed commit messages and documentation for any new features or bug fixes.


## License

SightSpeak Reader is licensed under the MIT License. See the LICENSE file for more details.

### Acknowledgments

**BS_thread_pool Library:** This project makes use of the BS_thread_pool library by Burak Sezer. The BS_thread_pool library is licensed under the MIT License. You can find the license information [here](https://github.com/bshoshany/thread-pool/blob/master/LICENSE).

## Feedback and Support
As this is a beta release, we value your feedback to help improve the project. Please submit any issues, feature requests, or general feedback via the GitHub Issues page.

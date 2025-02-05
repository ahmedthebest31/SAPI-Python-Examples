# SAPI-Python-Examples

This repository contains Python code examples for using the Microsoft Speech API (SAPI) on Windows.  It's a companion to my YouTube video [https://www.youtube.com/watch?v=mxn4PD752_Q] where I explain the differences between SAPI 4 and SAPI 5, and how to use SAPI for text-to-speech (TTS) and speech recognition (SR) in Python.

## What is SAPI?

The Microsoft Speech API (SAPI) is a powerful technology that allows applications to interact with speech.  It provides a standard interface for speech recognition and text-to-speech functionalities.  This repository focuses on SAPI 5, the current and recommended version.

## SAPI 4 vs. SAPI 5

In the video, I discuss the key differences between SAPI 4 and SAPI 5.  SAPI 5 is a major redesign offering better performance, flexibility, and features.  **It's crucial to use SAPI 5 for modern Windows development.** SAPI 4 is outdated and not recommended.

## Python Examples

This repository provides the following Python examples:

*   **Text-to-Speech (TTS):**  Demonstrates how to use SAPI to convert text to speech, including voice selection, rate adjustment, and volume control.
*   **Speech Recognition (SR):** Shows how to use SAPI to recognize spoken words, including basic recognition and grammar-based recognition for improved accuracy.

## Code Usage

1.  **Clone the repository:** `git clone https://github.com/[your-username]/SAPI-Python-Examples.git`
2.  **Install pywin32:** `pip install pywin32`
3.  **Run the examples:** Execute the Python scripts (e.g., `python speech_test.py`).
4.  **Grammar File (for SR):** If using grammar-based recognition, create an XML file (e.g., `mygrammar.xml`) in the same directory as your script, and put the grammar XML content into it.

## SAPI Documentation and Downloads

*   **Microsoft Speech SDK 5.1:**  You'll need to download and install the core SDK and the appropriate language pack.  Search the Microsoft website for "Microsoft Speech SDK 5.1" to find the download links.  (Direct links can sometimes change, so searching is the most reliable method).
*   **SAPI Documentation:** Search for "Microsoft Speech API (SAPI) 5.x Documentation" on the Microsoft website for detailed information about the API.

## Contributing

Contributions are welcome!  Feel free to submit pull requests or open issues.

## License

mit License


## YouTube Video

[[https://www.youtube.com/watch?v=mxn4PD752_Q](https://www.youtube.com/watch?v=mxn4PD752_Q)]
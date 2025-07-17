
[**English**](./README.md) | [**中文**](README_cn.md)
# Document Desensitization Tool

## Tool Description

This tool performs a simple desensitization function by replacing keyword occurrences in the filenames of `.docx` and `.doc` files within a specified `folder_path`.

## Notice

1.  To process `.doc` files, the tool first converts them to the `.docx` format using LibreOffice. Therefore, LibreOffice must be installed on your local machine, and the `soffice` command must be executable from the command line.
2.  This tool will permanently delete the source files. It is crucial to **back up your original files** before use.

## Usage

To use this tool, you need to modify the following variables in the script:

*   `folder_path`: Specify the source and destination folder for the files.
*   `keywords_to_replace`: Provide a list of sensitive words to be replaced.
*   `replacement_text`: Define the text that will replace the sensitive keywords.

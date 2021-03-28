# sakai-comments

Utility script to export instructor comments from an Excel spreadsheet to individual `comments.txt` files for a downloaded Sakai assignment.

## Installation

```
pip install -r requirements.txt
```

## Usage

If comments have been added to a column with the header `Comment` to `grades.xls`:

```
File path for grades.xls? path/to/Assignment 1/grades.xls
Assignment path (the folder with each student's folder)? path/to/Assignment 1
Column header for comments? Comment
```

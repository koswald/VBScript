###### The `algorithm` folder

## ReadMe.md

Algorithm development details.

## Contents
[ParseArgs propery](#parseargs-propery)  
[Pattern property](#pattern-property)

# ParseArgs propery

The [HTAApp class] `ParseArgs` propery parses the mshta.exe command-line arguments. This is unnecessary for .vbs and .wsf files, because the WScript.Arguments object provides the arguments already separated, with any wapping quotes stripped off. The mshta.exe raw arguments typically include the full path of the .hta file, wrapped in double quotes.

[HTAApp class]: ../../class/HTAApp.vbs

## Requirements 

- Multiple arguments must be separated by spaces.
- Arguments with spaces *must* be wrapped with quotes (i.e. double quotes).
- Arguments without spaces *may* be wrapped with quotes.
- Quoted arguments may be mixed with unquoted arguments.
- Quotes are used only in pairs and only for helping to define where to separate one argument from another.
- Return an array of arguments with no quotes.

## Algorithm synopsis

Temporarily wrap quoteless arguments with quotes, then split the modified command line string into an array.

## Algorithm

- If there are no arguments, return a zero-element array.
- Read one character at a time from left to right. Characters fall into one of three categories: double-quotes, spaces, and everything else.
- If an odd number of quotes have been read, then a quote-wrapped argument is being read...
    - Raise an error if the quote doesn't have a space immediately to its left.
    - Raise an error if there are an odd number of double-quotes.
- If an even number of quotes have been read, then a quoteless argument is being read...
    - Raise an error if the trailing quote, if any, of the previous argument doesn't have a space immediately to the right.
    - Temporarily add a leading quote.
    - Temporarily add a trailing quote.
    - Remove multiple spaces between arguments.
- Remove the leading and trailing quotes, if any.
- Convert the rebuilt arguments string to an array.
## Examples

| Argument | result |
| --- | :---: |
| /folder:C:\myfolder | ok |
| "/folder:C:\my folder" | ok |
| /folder:"C:\my folder" | error |



# Pattern property

The [RegExFunctions class] `Pattern` property converts a wildcard expression to a regex pattern.

[RegExFunctions class]: ../../class/RegExFunctions.vbs

## Requirements

- The following characters, *Group 1*, are invalid in Windows&reg; file names, so handle them specially, ignore them, or raise an error:
    ```
    \ / : * ? | " < >
    ```
- The following characters, *Group 2*, are regex special characters, so if not already taken care of above, handle them specially or escape them with a backslash (`\`):
    ```
    ( ) . $ + [ ? \ ^ { |
    ```

## Algorithm

*Group 1* characters:
- Replace `*` with `.*`
- Replace `?` with `.{1}` Do this last so that the `{` does not get escaped.
- `|` is the regex delimiter, so ignore it.
- Raise an error on the remaining *Group 1* characters:
    ```
    \ / : " < >
    ```  

*Group 2* characters:
- Replace `.` with `\.` (escape using `\`). Do this first, even before *Group 1* replacements, because `.` is used in other replacements.  
- Escape (using `\`) the remaining *Group 2* characters:
    ```
    ( ) $ + [ ^ {
    ```



<br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
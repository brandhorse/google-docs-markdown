# Google Docs Markdown 

## Author
- Name: Jeff Barnaby - jeffbarnaby.co | Brandhorse | Stroelli LLC.
- Contact: info@brandhorse.com
- GitHub: https://github.com/brandhorse/google-docs-markdown

### Project Description
 Parses Markdown syntax in the given Google Doc and applies corresponding Google Docs formatting. This includes headings, bold and italic text, bullet lists, blockquotes, code blocks, images, tables, strikethrough, horizontal rules and hyperlinks.

 Due to API limitations, this script DOES NOT convert paragraphs to ListItem objects with automatic bullet or numbered formatting but prepares the text accordingly. For full list functionality, including automatic numbering and bullet points, manual adjustment in the Google Docs UI may be necessary after script execution or exploring more complex scripting solutions.

## License
This project is licensed under the GNU GENERAL PUBLIC LICENSE - see the (LICENSE.md) file for details.


## How to Use
Instructions for using the provided code in Google Apps Script in a Google Doc:

1. Open the Google Doc you want to add the script to.
2. From the menu, select Tools > Script editor. This will open a new script project for your doc.
3. Delete any code that is already in the script editor and paste in the provided code.
4. Make sure to update any placeholder values in the code with your own specific values. For example, update the spreadsheet ID to match your own spreadsheet.
5. Add any additional functions or logic you need to integrate the script into your doc.
6. Save the project by clicking the save icon. Name the project something relevant.
7. To run the script, you can either: Click the Run icon to run the whole script or Call specific functions from the script manually by adding them into your doc with a script tag like: <?= myFunction() ?>
8. Authorize the script when prompted to allow it to run and access your Google services/data.
9. Test the script by reloading the doc or running the functions again.
10. If needed, view the script logs to debug any issues by selecting View > Logs from the script editor menu.

Once working as expected, the script will now run each time the doc is opened or edited to pull in the latest data.

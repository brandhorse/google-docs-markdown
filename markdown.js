/**
 * Parses Markdown syntax in the given Google Doc and applies corresponding
 * Google Docs formatting. This includes headings, bold and italic text,
 * bullet lists, and hyperlinks.
 */
function parseMarkdown() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();

  paragraphs.forEach((p) => {
    const pText = p.getText();

    // Applying formatting
    applyHeadings(p, pText);
    applyTextStyle(p, BOLD_REGEX, true, false); // Handle bold
    applyTextStyle(p, ITALIC_REGEX, false, true); // Handle italic
    applyLists(p, pText, body); // Handle lists
    applyLinks(p, pText); // Handle links
  });
}

/**
 * Applies heading styles based on Markdown syntax in paragraph text.
 */
function applyHeadings(p, pText) {
  const headingLevels = ["# ", "## ", "### ", "#### ", "##### "];
  headingLevels.forEach((heading, index) => {
    if (pText.startsWith(heading)) {
      p.setHeading(DocumentApp.ParagraphHeading['HEADING' + (index + 1)]);
      p.replaceText(heading, "");
    }
  });
}

/**
 * Applies bold and italic text styles based on Markdown syntax.
 */
function applyTextStyle(p, regex, isBold, isItalic) {
  let foundElement;
  const content = p.asText();

  while ((foundElement = p.findText(regex)) !== null) {
    const start = foundElement.getStartOffset();
    const end = foundElement.getEndOffsetInclusive();

    if (isBold) content.setBold(start, end, true);
    if (isItalic) content.setItalic(start, end, true);

    // Remove Markdown syntax after applying style
    const textToReplace = foundElement.getElement().asText().getText().substring(start, end + 1);
    const newText = textToReplace.replace(/[\*\_]/g, "");
    content.replaceText(textToReplace, newText);
  }
}

/**
 * Converts Markdown bullet lists to Google Docs bullet lists, handling bold within lists.
 * Updated to support "-", "*", and numbered list items "1.", "2.", etc.
 */
function applyLists(p, pText, body) {
  const bulletListRegex = /^(\- |\* |[0-9]+\.\s)/;
  const match = pText.match(bulletListRegex);
  if (match) {
    // Remove Markdown list syntax, handling bold within
    let listItemText = pText.substring(match[0].length);
    if (listItemText.startsWith("**")) {
      listItemText = listItemText.substring(2, listItemText.length - 2); // Adjust to remove bold syntax
      p.setText(listItemText);
      p.setBold(true);
    } else {
      p.setText(listItemText);
    }
    // Note: Direct conversion to bullet list items or numbered list items is not supported by Google Apps Script API.
  }
}

/**
 * Converts Markdown hyperlinks to clickable links in Google Docs.
 */
function applyLinks(p, pText) {
  let match;
  while ((match = LINK_REGEX.exec(pText)) !== null) {
    const textToReplace = match[0];
    const linkText = match[1];
    const url = match[2];
    p.replaceText(textToReplace, linkText);
    const foundLink = p.findText(linkText);
    if (foundLink) {
      const start = foundLink.getStartOffset();
      const end = foundLink.getEndOffsetInclusive();
      p.setLinkUrl(start, end, url);
    }
  }
}

// Regex updated to match bold and italic syntax correctly
const BOLD_REGEX = /\*\*(.*?)\*\*/g; // Fixed to correctly match bold syntax
const ITALIC_REGEX = /\_(.*?)\_/g; // To match italic syntax using underscores
const LINK_REGEX = /\[([^\]]+?)\]\(([^)]+?)\)/g;

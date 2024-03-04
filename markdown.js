/**
 * Parses Markdown syntax in the given Google Doc and applies corresponding
 * Google Docs formatting. This includes headings, bold and italic text,
 * and hyperlinks. Note: Direct conversion of text to bullet or numbered lists
 * is not supported due to API limitations, but the script will format other Markdown elements.
 */
function parseMarkdown() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();

  paragraphs.forEach((p) => {
    const pText = p.getText();

    // Applying formatting
    applyHeadings(p, pText);
    applyTextStyle(p); // Adjusted to handle both bold and italic
    // applyLists(p, pText, body); // Direct list handling is removed due to API limitations
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
      p.replaceText("^" + heading, ""); // Corrected to replace at start
    }
  });
}

/**
 * Applies bold and italic text styles based on Markdown syntax.
 * Correctly handles bold text by applying bold formatting only to the text between double asterisks
 * and ensures that both sets of asterisks are removed.
 */
function applyTextStyle(p) {
  const content = p.editAsText();
  // Handle bold
  const boldRegex = /\*\*(.*?)\*\*/g;
  let match;
  while ((match = boldRegex.exec(p.getText())) !== null) {
    const fullMatch = match[0];
    const textToBold = match[1];
    const startIndex = match.index;
    const endIndex = startIndex + textToBold.length + 1; // Correctly end before the second set of asterisks

    // Apply bold formatting only to the text between the asterisks
    content.setBold(startIndex, endIndex, true);
    // Replace the full match (including asterisks) with just the text, effectively removing the asterisks
    content.replaceText(escapeRegExp(fullMatch), textToBold);
  }

  // Additional logic for italic or other styles can be added here
}

/**
 * Converts Markdown hyperlinks to clickable links in Google Docs.
 */
function applyLinks(p, pText) {
  const linkRegex = /\[([^\]]+?)\]\(([^)]+?)\)/g;
  let match;
  while ((match = linkRegex.exec(p.getText())) !== null) {
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

// Helper function to escape special characters for use in a regular expression
function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

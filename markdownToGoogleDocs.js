/**
 * Parses Markdown syntax in the given Google Doc and applies corresponding
 * Google Docs formatting. This includes headings, bold and italic text,
 * code blocks, and hyperlinks.
 */
function parseMarkdown() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();
  let inCodeBlock = false;

  for (let i = 0; i < paragraphs.length; i++) {
    let p = paragraphs[i];
    let pText = p.getText().trim();

    // Toggle the inCodeBlock flag and clear markers
    if (pText.match(/^```(\w+)?/)) {
      inCodeBlock = !inCodeBlock; // Toggle state
      p.clear(); // Remove the marker
      if (!inCodeBlock) {
        // If we've just closed a code block, move to the next paragraph without further processing
        continue;
      }
    } else if (inCodeBlock) {
      // If we are inside a code block, format the paragraph as code
      p.editAsText().setFontFamily('Courier New'); // Apply monospace font
    } else {
      // Applying other markdown formatting outside of code blocks
      applyHeadings(p, p.getText());
      applyTextStyle(p); // Handle both bold and italic
      applyLinks(p, p.getText()); // Handle links
    }
  }
}

/**
 * Applies heading styles based on Markdown syntax in paragraph text.
 */
function applyHeadings(p, pText) {
  const headingLevels = ["# ", "## ", "### ", "#### ", "##### "];
  headingLevels.forEach((heading, index) => {
    if (pText.startsWith(heading)) {
      p.setHeading(DocumentApp.ParagraphHeading['HEADING' + (index + 1)]);
      p.replaceText(heading, ""); // Remove markdown syntax for heading
    }
  });
}

/**
 * Applies bold and italic text styles based on Markdown syntax.
 */
function applyTextStyle(p) {
  const content = p.editAsText();
  // Handle bold
  const boldRegex = /\*\*(.*?)\*\*/g;
  let match;
  while ((match = boldRegex.exec(p.getText())) !== null) {
    const textToBold = match[1];
    const startIndex = match.index;
    const endIndex = startIndex + textToBold.length + 3; // Adjust for the length of bold syntax

    // Apply bold formatting only to the text between the asterisks
    content.setBold(startIndex, endIndex, true);
    // Remove the asterisks
    content.replaceText(/\*\*(.*?)\*\*/, "$1");
  }

  // Implement similar logic for italics if needed
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

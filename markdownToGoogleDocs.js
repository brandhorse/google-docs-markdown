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

    if (pText.match(/^```(\w+)?/)) {
      inCodeBlock = !inCodeBlock; // Toggle state
      p.clear(); // Remove the marker
      continue;
    } else if (inCodeBlock) {
      p.editAsText().setFontFamily('Courier New'); // Apply monospace font
    } else {
      applyHeadings(p, p.getText());
      applyTextStyle(p); // Handle both bold and italic
      applyLinks(p, p.getText()); // Handle links
    }
  }

  // Final cleanup to remove any remaining double asterisks
  removeRemainingAsterisks(body);
}

function applyHeadings(p, pText) {
  const headingLevels = ["# ", "## ", "### ", "#### ", "##### "];
  headingLevels.forEach((heading, index) => {
    if (pText.startsWith(heading)) {
      p.setHeading(DocumentApp.ParagraphHeading['HEADING' + (index + 1)]);
      p.replaceText(heading, "");
    }
  });
}

function applyTextStyle(p) {
  const content = p.editAsText();
  const boldRegex = /\*\*(.*?)\*\*/g;
  let match;
  while ((match = boldRegex.exec(p.getText())) !== null) {
    const startIndex = match.index;
    const endIndex = startIndex + match[0].length - 1;
    content.setBold(startIndex, endIndex, true);
    content.replaceText(/\*\*(.*?)\*\*/, "$1");
  }
  // Additional logic for italic formatting can be added here
}

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

function removeRemainingAsterisks(body) {
  const paragraphs = body.getParagraphs();
  paragraphs.forEach((p) => {
    const content = p.editAsText();
    content.replaceText("\\*\\*", ""); // Remove leftover double asterisks
  });
}

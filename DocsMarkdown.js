/**
 * Parses Markdown syntax in the given Google Doc and applies corresponding
 * Google Docs formatting. This includes headings, bold and italic text,
 * bullet lists, blockquotes, code blocks, images, tables, strikethrough, horizontal rules
 * and hyperlinks.
 */
function parseMarkdown() {
  const doc = DocumentApp.getActiveDocument();

  /**
   * Gets all body paragraphs in the document.
   * 
   * @returns {Paragraph[]} Array of Paragraph objects
   */
  const docParagraphs = doc.getBody().getParagraphs();

  docParagraphs.forEach(applyMarkdownFormatting);
}

/**
 * Applies Markdown formatting to a paragraph.
 * 
 * @param {Paragraph} paragraph - Paragraph to format
 */
function applyMarkdownFormatting(paragraph) {
  const paragraphText = paragraph.getText();

  // Applying formatting
  applyHeadings(paragraph, paragraphText);
  applyTextStyle(paragraph, BOLD_REGEX, true, false);
  applyTextStyle(paragraph, ITALIC_REGEX, false, true);
  applyLists(paragraph, paragraphText, doc.getBody());
  applyBlockquotes(paragraph, paragraphText);
  applyCodeBlocks(paragraph, paragraphText);
  applyImages(paragraph, paragraphText);
  applyTables(paragraph, paragraphText);
  applyStrikethrough(paragraph, paragraphText);
  applyHorizontalRules(paragraph, paragraphText);
  applyLinks(paragraph, paragraphText);
}

/**
 * Converts Markdown blockquotes to indented paragraphs in Google Docs.
 * 
 * @param {Paragraph} paragraph - Paragraph to format 
 * @param {string} text - Paragraph text
 */
function applyBlockquotes(paragraph, text) {
  if (text.startsWith("> ")) {
    // Indent paragraph to simulate blockquote
    const indent = {};
    indent[DocumentApp.Attribute.INDENT_START] = 72;
    paragraph.setAttributes(indent);

    // Remove > character
    paragraph.replaceText("> ", "");
  }
}

/**
 * Applies background shading for Markdown code blocks.
 * 
 * @param {Paragraph} paragraph - Paragraph containing code block
 * @param {string} text - Paragraph text
 */
function applyCodeBlocks(paragraph, text) {
  if (text.startsWith("```") && text.endsWith("```")) {
    const start = paragraph.findText("```").getStartOffset();
    const end = paragraph.findText("```", start + 1).getEndOffsetInclusive();

    const background = {};
    background[DocumentApp.Attribute.BACKGROUND_COLOR] = "#e8e8e8";
    paragraph.setAttributes(background, start, end);

    // Remove ``` characters
    paragraph.deleteText(start, 3);
    paragraph.deleteText(end - 3, end);
  }
}

/**
 * Inserts images based on Markdown image syntax.
 * 
 * @param {Paragraph} paragraph - Paragraph containing image syntax
 * @param {string} text - Paragraph text 
 */
function applyImages(paragraph, text) {
  let match;
  while ((match = IMAGE_REGEX.exec(text)) !== null) {
    const altText = match[1];
    const imageUrl = match[2];

    // Validate inputs
    if (!altText || !imageUrl) {
      continue;
    }

    // Insert image
    const img = paragraph.insertImage(imageUrl);
    img.setAltDescription(altText);

    // Remove Markdown syntax
    paragraph.deleteText(match.index, match[0].length);
  }
}

/**
 * Converts Markdown tables to Google Docs tables.
 * 
 * @param {Paragraph} paragraph - Paragraph containing Markdown table syntax
 * @param {string} text - Paragraph text
 */
function applyTables(paragraph, text) {
  // TODO: Implement table parsing and insertion
}

/**
 * Applies strikethrough formatting based on Markdown syntax.
 * 
 * @param {Paragraph} paragraph - Paragraph containing strikethrough syntax
 * @param {string} text - Paragraph text
 */
function applyStrikethrough(paragraph, text) {
  let foundElement;
  const content = paragraph.asText();

  while ((foundElement = paragraph.findText(STRIKETHROUGH_REGEX)) !== null) {
    const start = foundElement.getStartOffset();
    const end = foundElement.getEndOffsetInclusive();

    content.setStrikethrough(start, end, true);

    // Remove Markdown syntax
    const textToReplace = foundElement.getElement().asText().getText().substring(start, end +
/**
 * Renders the small subset of Markdown the summarizer actually produces
 * (## headers, - / * bullet lists, plain paragraphs). Deliberately not a
 * full Markdown library - the summarizer's prompts are constrained to this
 * exact shape, so a tiny purpose-built renderer avoids an extra dependency
 * for something this narrow.
 */
import * as React from 'react';

interface IBlock {
  type: 'heading' | 'list' | 'paragraph';
  content: string[];
}

function parseIntoBlocks(markdown: string): IBlock[] {
  const lines = markdown.split('\n');
  const blocks: IBlock[] = [];
  let currentList: string[] | undefined;

  const flushList = (): void => {
    if (currentList && currentList.length > 0) {
      blocks.push({ type: 'list', content: currentList });
    }
    currentList = undefined;
  };

  for (const rawLine of lines) {
    const line = rawLine.trim();

    if (line.length === 0) {
      flushList();
      continue;
    }

    if (/^#{1,6}\s+/.test(line)) {
      flushList();
      blocks.push({ type: 'heading', content: [line.replace(/^#{1,6}\s+/, '')] });
      continue;
    }

    if (/^[-*]\s+/.test(line)) {
      const item = line.replace(/^[-*]\s+/, '');
      currentList = currentList || [];
      currentList.push(item);
      continue;
    }

    flushList();
    blocks.push({ type: 'paragraph', content: [line] });
  }

  flushList();
  return blocks;
}

export interface ISummaryMarkdownProps {
  markdown: string;
}

export const SummaryMarkdown: React.FC<ISummaryMarkdownProps> = ({ markdown }) => {
  const blocks = parseIntoBlocks(markdown);

  return (
    <>
      {blocks.map((block, index) => {
        if (block.type === 'heading') {
          return (
            <h3 key={index} style={{ fontSize: 15, fontWeight: 600, margin: '14px 0 6px' }}>
              {block.content[0]}
            </h3>
          );
        }
        if (block.type === 'list') {
          return (
            <ul key={index} style={{ margin: '4px 0 10px', paddingLeft: 20 }}>
              {block.content.map((item, itemIndex) => (
                <li key={itemIndex} style={{ fontSize: 13, lineHeight: 1.5, marginBottom: 4 }}>
                  {item}
                </li>
              ))}
            </ul>
          );
        }
        return (
          <p key={index} style={{ fontSize: 13, lineHeight: 1.5, margin: '4px 0 10px' }}>
            {block.content[0]}
          </p>
        );
      })}
    </>
  );
};

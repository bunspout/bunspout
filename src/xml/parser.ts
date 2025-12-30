import * as sax from 'sax';
import type { XmlEvent } from '../types';

/**
 * Parses XML bytes into XmlEvent stream
 */
export async function* parseXmlEvents(
  bytes: AsyncIterable<Uint8Array>,
): AsyncIterable<XmlEvent> {
  const parser = sax.parser(true, {
    // strict: true (first parameter) - Note: despite name, this allows fragments
    trim: false,
    normalize: false,
    lowercase: false,
    xmlns: false,
    position: false,
  });

  const eventQueue: XmlEvent[] = [];
  let error: Error | null = null;
  let finished = false;
  let resolveNext: (() => void) | null = null;

  parser.onopentag = (node: sax.Tag | sax.QualifiedTag) => {
    const attributes: Record<string, string> = {};
    if (node.attributes) {
      for (const key in node.attributes) {
        const attr = node.attributes[key];
        if (typeof attr === 'string') {
          attributes[key] = attr;
        } else if (attr && typeof attr === 'object' && 'value' in attr) {
          attributes[key] = String(attr.value);
        }
      }
    }
    eventQueue.push({
      type: 'startElement',
      name: node.name,
      attributes,
    });
    if (resolveNext) {
      resolveNext();
      resolveNext = null;
    }
  };

  parser.onclosetag = (name: string) => {
    eventQueue.push({
      type: 'endElement',
      name,
    });
    if (resolveNext) {
      resolveNext();
      resolveNext = null;
    }
  };

  parser.ontext = (text: string) => {
    if (text.trim()) {
      eventQueue.push({
        type: 'text',
        text,
      });
      if (resolveNext) {
        resolveNext();
        resolveNext = null;
      }
    }
  };

  parser.onerror = (err: Error) => {
    error = err;
    finished = true;
    if (resolveNext) {
      resolveNext();
      resolveNext = null;
    }
  };

  parser.onend = () => {
    finished = true;
    if (resolveNext) {
      resolveNext();
      resolveNext = null;
    }
  };

  // Start consuming bytes in background
  const consumePromise = (async () => {
    try {
      for await (const chunk of bytes) {
        const text = new TextDecoder().decode(chunk);
        parser.write(text);
      }
      parser.close();
    } catch (err) {
      parser.onerror(err as Error);
    }
  })();

  // Yield events as they come
  while (!finished || eventQueue.length > 0) {
    if (error) {
      throw error;
    }
    if (eventQueue.length > 0) {
      yield eventQueue.shift()!;
    } else if (!finished) {
      // Wait for next event
      await new Promise<void>((resolve) => {
        resolveNext = resolve;
      });
    } else {
      break;
    }
  }

  // Wait for consumption to complete
  await consumePromise;
}

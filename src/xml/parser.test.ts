import { describe, test, expect } from 'bun:test';
import { parseXmlEvents } from './parser';
import type { XmlEvent } from '../types';

describe('XML Parser', () => {
  test('should parse simple element → startElement event', async () => {
    const bytes = async function* () {
      yield new TextEncoder().encode('<row></row>');
    }();

    const events: XmlEvent[] = [];
    for await (const event of parseXmlEvents(bytes)) {
      events.push(event);
    }

    expect(events.length).toBeGreaterThan(0);
    const startEvent = events.find((e) => e.type === 'startElement' && e.name === 'row');
    expect(startEvent).toBeDefined();
    expect(startEvent?.type).toBe('startElement');
    expect(startEvent?.name).toBe('row');
  });

  test('should parse element with attributes → startElement with attributes', async () => {
    const bytes = async function* () {
      yield new TextEncoder().encode('<row r="1" spans="1:2"></row>');
    }();

    const events: XmlEvent[] = [];
    for await (const event of parseXmlEvents(bytes)) {
      events.push(event);
    }

    const startEvent = events.find(
      (e) => e.type === 'startElement' && e.name === 'row',
    );
    expect(startEvent).toBeDefined();
    expect(startEvent?.attributes).toEqual({ r: '1', spans: '1:2' });
  });

  test('should parse text content → text event', async () => {
    const bytes = async function* () {
      yield new TextEncoder().encode('<c>Hello</c>');
    }();

    const events: XmlEvent[] = [];
    for await (const event of parseXmlEvents(bytes)) {
      events.push(event);
    }

    const textEvent = events.find((e) => e.type === 'text');
    expect(textEvent).toBeDefined();
    expect(textEvent?.type).toBe('text');
    expect(textEvent?.text).toBe('Hello');
  });

  test('should parse nested elements → sequence of events', async () => {
    const bytes = async function* () {
      yield new TextEncoder().encode('<row><c>1</c></row>');
    }();

    const events: XmlEvent[] = [];
    for await (const event of parseXmlEvents(bytes)) {
      events.push(event);
    }

    const rowStart = events.find((e) => e.type === 'startElement' && e.name === 'row');
    const cStart = events.find((e) => e.type === 'startElement' && e.name === 'c');
    const textEvent = events.find((e) => e.type === 'text');
    const cEnd = events.find((e) => e.type === 'endElement' && e.name === 'c');
    const rowEnd = events.find((e) => e.type === 'endElement' && e.name === 'row');

    expect(rowStart).toBeDefined();
    expect(cStart).toBeDefined();
    expect(textEvent).toBeDefined();
    expect(textEvent?.text).toBe('1');
    expect(cEnd).toBeDefined();
    expect(rowEnd).toBeDefined();
  });

  test('should handle XML errors gracefully', async () => {
    const bytes = async function* () {
      yield new TextEncoder().encode('<row><unclosed>');
    }();

    let errorCaught = false;
    try {
      const events: XmlEvent[] = [];
      for await (const event of parseXmlEvents(bytes)) {
        events.push(event);
      }
    } catch (error) {
      errorCaught = true;
      expect(error).toBeDefined();
    }
    // XML parser should either throw or handle gracefully
    expect(errorCaught).toBe(true);
  });

  test('should handle chunked input', async () => {
    const encoder = new TextEncoder();
    const bytes = async function* () {
      yield encoder.encode('<row>');
      yield encoder.encode('<c>');
      yield encoder.encode('Test');
      yield encoder.encode('</c>');
      yield encoder.encode('</row>');
    }();

    const events: XmlEvent[] = [];
    for await (const event of parseXmlEvents(bytes)) {
      events.push(event);
    }

    const textEvent = events.find((e) => e.type === 'text');
    expect(textEvent?.text).toBe('Test');
  });
});

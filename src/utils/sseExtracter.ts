import { createParser, EventSourceMessage } from 'eventsource-parser';

export function getTextAsSseMessage(chunk: string, handler: (event?: EventSourceMessage | null, error?: Error | null) => void): void {
    try {
        let parser = createParser({
            onEvent(event) {
                handler(event, null);
            },
            onError(error) {
                handler(null, error);
            },
        });
        parser.feed(chunk);
    } catch (error) {
        console.error('Error parsing JSON:', error);
    }
}
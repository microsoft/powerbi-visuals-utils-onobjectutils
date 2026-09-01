import { vi } from "vitest";

import {
    HtmlSubSelectableClass,
    SubSelectableDisplayNameAttribute,
    SubSelectableObjectNameAttribute,
    SubSelectableTypeAttribute,
} from "../src/HtmlSubSelectionHelper";

export interface Rect {
    x?: number;
    y?: number;
    width: number;
    height: number;
}

// Laid out for real instead of stubbing `getBoundingClientRect`: fixed positioning makes the box
// viewport-relative, so the element's rect is exactly the one requested here.
export function setRect(element: Element, rect: Rect): void {
    const x = rect.x ?? 0;
    const y = rect.y ?? 0;
    const style = (element as HTMLElement).style;
    style.position = "fixed";
    style.boxSizing = "border-box";
    style.margin = "0";
    style.padding = "0";
    style.border = "0";
    style.left = `${x}px`;
    style.top = `${y}px`;
    style.width = `${rect.width}px`;
    style.height = `${rect.height}px`;
}

export interface SubSelectableOptions {
    objectName: string;
    displayName?: string;
    subSelectionType?: number;
    rect?: Rect;
    parent?: HTMLElement;
    attributes?: Record<string, string>;
}

export function createSubSelectable(host: HTMLElement, options: SubSelectableOptions): HTMLElement {
    const element = document.createElement("div");
    element.classList.add(HtmlSubSelectableClass);
    element.setAttribute(SubSelectableObjectNameAttribute, options.objectName);
    if (options.displayName !== undefined) {
        element.setAttribute(SubSelectableDisplayNameAttribute, options.displayName);
    }
    if (options.subSelectionType !== undefined) {
        element.setAttribute(SubSelectableTypeAttribute, `${options.subSelectionType}`);
    }
    for (const [name, value] of Object.entries(options.attributes ?? {})) {
        element.setAttribute(name, value);
    }

    setRect(element, options.rect ?? { x: 0, y: 0, width: 10, height: 10 });
    (options.parent ?? host).appendChild(element);
    return element;
}

export interface SubSelectionServiceMock {
    subSelect: ReturnType<typeof vi.fn>;
    updateRegionOutlines: ReturnType<typeof vi.fn>;
}

export function createSubSelectionServiceMock(): SubSelectionServiceMock {
    return {
        subSelect: vi.fn(),
        updateRegionOutlines: vi.fn(),
    };
}

/** Minimal ISelectionId stand-in: two ids are equal when they share the same key. */
export function createSelectionId(key: string, hasSelector: boolean = true): powerbi.visuals.ISelectionId {
    const selectionId = {
        getKey: (): string => key,
        getSelector: () => (hasSelector ? { metadata: key } : undefined),
        getSelectorsByColumn: () => undefined,
        hasIdentity: (): boolean => hasSelector,
        includes: (other: { getKey(): string }): boolean => other?.getKey() === key,
        equals: (other: { getKey(): string }): boolean => other?.getKey() === key,
    };
    return selectionId as unknown as powerbi.visuals.ISelectionId;
}

export function dispatch(element: Element, type: string, init: MouseEventInit = {}): MouseEvent {
    const event = new MouseEvent(type, { bubbles: true, cancelable: true, ...init });
    element.dispatchEvent(event);
    return event;
}

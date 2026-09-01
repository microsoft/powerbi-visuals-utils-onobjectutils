import { afterEach, describe, expect, it, vi } from "vitest";

import {
    debounce,
    equalsSelectionId,
    getObjectValues,
    getUniques,
    groupArrayElements,
    isArrayEmpty,
    isEqual,
} from "../src/helperFunctions";
import { createSelectionId } from "./testUtils";

describe("getObjectValues", () => {
    it("returns the values of the object", () => {
        expect(getObjectValues({ a: 1, b: 2, c: 3 })).toEqual([1, 2, 3]);
    });

    it("returns an empty array for an empty object", () => {
        expect(getObjectValues({})).toEqual([]);
    });

    it("returns an empty array when the object is not defined", () => {
        expect(getObjectValues(undefined as unknown as Record<string, number>)).toEqual([]);
        expect(getObjectValues(null as unknown as Record<string, number>)).toEqual([]);
    });
});

describe("groupArrayElements", () => {
    const createElement = (id: string): HTMLElement => {
        const element = document.createElement("div");
        element.id = id;
        return element;
    };

    it("groups elements by the key returned from the callback", () => {
        const first = createElement("first");
        const second = createElement("second");
        const third = createElement("third");

        const groups = groupArrayElements(
            [first, second, third],
            element => (element.id === "third" ? "b" : "a") as powerbi.visuals.SubSelectionRegionOutlineId,
        );

        expect(Object.keys(groups)).toEqual(["a", "b"]);
        expect(groups["a"]).toEqual([first, second]);
        expect(groups["b"]).toEqual([third]);
    });

    it("returns an empty object when the array is not defined", () => {
        expect(groupArrayElements(undefined as unknown as HTMLElement[], () => "a" as powerbi.visuals.SubSelectionRegionOutlineId)).toEqual({});
    });

    it("returns an empty object for an empty array", () => {
        expect(groupArrayElements([], () => "a" as powerbi.visuals.SubSelectionRegionOutlineId)).toEqual({});
    });
});

describe("isEqual", () => {
    it("returns true for identical primitives", () => {
        expect(isEqual(1, 1)).toBe(true);
        expect(isEqual("a", "a")).toBe(true);
        expect(isEqual(null, null)).toBe(true);
        expect(isEqual(undefined, undefined)).toBe(true);
    });

    it("returns false for different primitives", () => {
        expect(isEqual(1, 2)).toBe(false);
        expect(isEqual("a", "b")).toBe(false);
        expect(isEqual(null, undefined)).toBe(false);
    });

    it("returns true for deeply equal objects", () => {
        expect(isEqual({ a: 1, b: { c: [1, 2] } }, { a: 1, b: { c: [1, 2] } })).toBe(true);
    });

    it("returns false when the number of properties differs", () => {
        expect(isEqual({ a: 1 }, { a: 1, b: 2 })).toBe(false);
    });

    it("returns false when a nested value differs", () => {
        expect(isEqual({ a: 1, b: { c: 1 } }, { a: 1, b: { c: 2 } })).toBe(false);
    });

    it("returns false when only one of the values is an object", () => {
        expect(isEqual({ a: 1 }, "a")).toBe(false);
    });
});

describe("isArrayEmpty", () => {
    it("returns true for empty or missing arrays", () => {
        expect(isArrayEmpty([])).toBe(true);
        expect(isArrayEmpty(undefined as unknown as unknown[])).toBe(true);
        expect(isArrayEmpty(null as unknown as unknown[])).toBe(true);
    });

    it("returns false for a non empty array", () => {
        expect(isArrayEmpty([1])).toBe(false);
    });
});

describe("getUniques", () => {
    it("keeps only the first element of each equal group", () => {
        const array = [
            { id: 1, name: "a" },
            { id: 2, name: "b" },
            { id: 3, name: "a" },
        ];

        expect(getUniques(array, (a, b) => a.name === b.name)).toEqual([
            { id: 1, name: "a" },
            { id: 2, name: "b" },
        ]);
    });

    it("returns an empty array when the array is not defined", () => {
        expect(getUniques(undefined as unknown as number[], (a, b) => a === b)).toEqual([]);
    });
});

describe("debounce", () => {
    afterEach(() => {
        vi.useRealTimers();
    });

    it("invokes the function only once after the delay", () => {
        vi.useFakeTimers();
        const callback = vi.fn();
        const debounced = debounce(callback, 100);

        debounced();
        debounced();
        debounced();

        expect(callback).not.toHaveBeenCalled();
        vi.advanceTimersByTime(100);
        expect(callback).toHaveBeenCalledTimes(1);
    });

    it("passes the arguments and the calling context to the function", () => {
        vi.useFakeTimers();
        const callback = vi.fn();
        const context = { debounced: debounce(callback, 50) };

        context.debounced(1, "two");
        vi.advanceTimersByTime(50);

        expect(callback).toHaveBeenCalledWith(1, "two");
        expect(callback.mock.instances[0]).toBe(context);
    });
});

describe("equalsSelectionId", () => {
    it("returns true when both selection ids are missing", () => {
        expect(equalsSelectionId(undefined as unknown as powerbi.visuals.ISelectionId, null as unknown as powerbi.visuals.ISelectionId)).toBe(true);
    });

    it("returns false when only one selection id is defined", () => {
        expect(equalsSelectionId(createSelectionId("a"), undefined as unknown as powerbi.visuals.ISelectionId)).toBe(false);
        expect(equalsSelectionId(undefined as unknown as powerbi.visuals.ISelectionId, createSelectionId("a"))).toBe(false);
    });

    it("returns true for equal selection ids", () => {
        expect(equalsSelectionId(createSelectionId("a"), createSelectionId("a"))).toBe(true);
    });

    it("returns false for different selection ids", () => {
        expect(equalsSelectionId(createSelectionId("a"), createSelectionId("b"))).toBe(false);
    });

    it("returns true for the same instance", () => {
        const selectionId = createSelectionId("a");
        expect(equalsSelectionId(selectionId, selectionId)).toBe(true);
    });
});

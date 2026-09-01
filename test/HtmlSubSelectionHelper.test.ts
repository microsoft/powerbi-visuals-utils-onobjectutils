import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import {
    HtmlSubSelectableClass,
    HtmlSubSelectionHelper,
    SubSelectableAltObjectNameAttribute,
    SubSelectableDirectEdit,
    SubSelectableHideOutlineAttribute,
    SubSelectableObjectNameAttribute,
    SubSelectableRestrictingElementAttribute,
    SubSelectableSubSelectedAttribute,
} from "../src/HtmlSubSelectionHelper";
import { HtmlSubselectionHelperArgs } from "../src/types";
import {
    SubSelectionOutlineRestrictionType,
    SubSelectionOutlineType,
    SubSelectionOutlineVisibility,
    SubSelectionStylesType,
} from "./mocks/powerbiApiMock";
import {
    createSelectionId,
    createSubSelectable,
    createSubSelectionServiceMock,
    dispatch,
    setRect,
    SubSelectionServiceMock,
} from "./testUtils";

describe("HtmlSubSelectionHelper", () => {
    let host: HTMLElement;
    let subSelectionService: SubSelectionServiceMock;

    const createHelper = (args: Partial<HtmlSubselectionHelperArgs> = {}): HtmlSubSelectionHelper =>
        HtmlSubSelectionHelper.createHtmlSubselectionHelper({
            hostElement: host,
            subSelectionService: subSelectionService as unknown as powerbi.extensibility.IVisualSubSelectionService,
            ...args,
        });

    beforeEach(() => {
        document.body.replaceChildren();
        host = document.createElement("div");
        setRect(host, { x: 0, y: 0, width: 500, height: 500 });
        document.body.appendChild(host);
        subSelectionService = createSubSelectionServiceMock();
    });

    afterEach(() => {
        vi.useRealTimers();
    });

    describe("createHtmlSubselectionHelper", () => {
        it("marks the host element so the helper can recognize its own elements", () => {
            createHelper();

            expect(host.getAttribute("data-helper-host")).toBe("true");
        });
    });

    describe("element data", () => {
        it("round trips the sub-selection data of an element", () => {
            const element = document.createElement("div");
            const data = { outlineRestrictionOptions: { margin: { top: 1, left: 2, right: 3, bottom: 4 } } };

            HtmlSubSelectionHelper.setDataForElement(element, data);

            expect(HtmlSubSelectionHelper.getDataForElement(element)).toEqual(data);
        });

        it("returns null when the element has no sub-selection data", () => {
            expect(HtmlSubSelectionHelper.getDataForElement(document.createElement("div"))).toBeNull();
        });
    });

    describe("updateElementOutlines", () => {
        it("creates a group outline containing the rectangle of the element", () => {
            const helper = createHelper();
            const element = createSubSelectable(host, { objectName: "text", rect: { x: 10, y: 20, width: 30, height: 40 } });

            const [id] = helper.updateElementOutlines([element], SubSelectionOutlineVisibility.Active);

            expect(id).toBe("text");
            expect(helper.getRegionOutline(id)).toEqual({
                id: "text",
                visibility: SubSelectionOutlineVisibility.Active,
                outline: {
                    type: SubSelectionOutlineType.Group,
                    outlines: [{ type: SubSelectionOutlineType.Rectangle, x: 10, y: 20, width: 30, height: 40 }],
                },
            });
        });

        it("groups elements that share the same object name into a single region", () => {
            const helper = createHelper();
            const first = createSubSelectable(host, { objectName: "text", rect: { x: 0, y: 0, width: 10, height: 10 } });
            const second = createSubSelectable(host, { objectName: "text", rect: { x: 20, y: 0, width: 10, height: 10 } });
            const other = createSubSelectable(host, { objectName: "shape", rect: { x: 0, y: 40, width: 10, height: 10 } });

            const ids = helper.updateElementOutlines([first, second, other], SubSelectionOutlineVisibility.Active);

            expect(ids).toEqual(["text", "shape"]);
            expect((helper.getRegionOutline("text" as powerbi.visuals.SubSelectionRegionOutlineId)!.outline as powerbi.visuals.GroupSubSelectionOutline).outlines).toHaveLength(2);
            expect((helper.getRegionOutline("shape" as powerbi.visuals.SubSelectionRegionOutlineId)!.outline as powerbi.visuals.GroupSubSelectionOutline).outlines).toHaveLength(1);
        });

        it("appends the selection key to the region id when a selection id callback is provided", () => {
            const element = createSubSelectable(host, { objectName: "text" });
            const helper = createHelper({ selectionIdCallback: () => createSelectionId("key1") });

            expect(helper.updateElementOutline(element, SubSelectionOutlineVisibility.Active)).toBe("text___key1");
        });

        it("keeps the plain object name when the selection id has no selector", () => {
            const element = createSubSelectable(host, { objectName: "text" });
            const helper = createHelper({ selectionIdCallback: () => createSelectionId("key1", false /* hasSelector */) });

            expect(helper.updateElementOutline(element, SubSelectionOutlineVisibility.Active)).toBe("text");
        });

        it("skips elements without a visible area", () => {
            const helper = createHelper();
            const element = createSubSelectable(host, { objectName: "text", rect: { width: 0, height: 0 } });

            const id = helper.updateElementOutline(element, SubSelectionOutlineVisibility.Active);

            expect((helper.getRegionOutline(id)!.outline as powerbi.visuals.GroupSubSelectionOutline).outlines).toEqual([]);
        });

        it("reads the direct edit definition from the element", () => {
            const helper = createHelper();
            const directEdit = { reference: { objectName: "text", propertyName: "value" }, style: 3 };
            const element = createSubSelectable(host, {
                objectName: "text",
                attributes: { [SubSelectableDirectEdit]: JSON.stringify(directEdit) },
            });

            const id = helper.updateElementOutline(element, SubSelectionOutlineVisibility.Active);
            const outlines = (helper.getRegionOutline(id)!.outline as powerbi.visuals.GroupSubSelectionOutline).outlines;

            expect((outlines[0] as powerbi.visuals.RectangleSubSelectionOutline).cVDirectEdit).toEqual(directEdit);
        });

        it("renders the outlines through the sub-selection service unless render is suppressed", () => {
            const helper = createHelper();
            const element = createSubSelectable(host, { objectName: "text" });

            helper.updateElementOutline(element, SubSelectionOutlineVisibility.Active, true /* suppressRender */);
            expect(subSelectionService.updateRegionOutlines).not.toHaveBeenCalled();

            helper.updateElementOutline(element, SubSelectionOutlineVisibility.Active);
            expect(subSelectionService.updateRegionOutlines).toHaveBeenCalledTimes(1);
            expect(subSelectionService.updateRegionOutlines).toHaveBeenCalledWith([helper.getRegionOutline("text" as powerbi.visuals.SubSelectionRegionOutlineId)]);
        });
    });

    describe("outline restrictions", () => {
        const createRestrictedElement = (type: number): HTMLElement => {
            const restrictingElement = document.createElement("div");
            restrictingElement.setAttribute(SubSelectableRestrictingElementAttribute, `${type}`);
            setRect(restrictingElement, { x: 10, y: 10, width: 50, height: 50 });
            host.appendChild(restrictingElement);

            return createSubSelectable(host, {
                objectName: "text",
                rect: { x: 0, y: 0, width: 100, height: 100 },
                parent: restrictingElement,
            });
        };

        it("clamps the outline to the restricting element", () => {
            const helper = createHelper();
            const element = createRestrictedElement(SubSelectionOutlineRestrictionType.Clamp);

            const id = helper.updateElementOutline(element, SubSelectionOutlineVisibility.Active);
            const outlines = (helper.getRegionOutline(id)!.outline as powerbi.visuals.GroupSubSelectionOutline).outlines;

            expect(outlines[0]).toMatchObject({ x: 10, y: 10, width: 50, height: 50 });
        });

        it("applies the margin of the sub-selection data while clamping", () => {
            const helper = createHelper();
            const element = createRestrictedElement(SubSelectionOutlineRestrictionType.Clamp);
            HtmlSubSelectionHelper.setDataForElement(element, {
                outlineRestrictionOptions: { margin: { top: 5, left: 5, right: 5, bottom: 5 } },
            });

            const id = helper.updateElementOutline(element, SubSelectionOutlineVisibility.Active);
            const outlines = (helper.getRegionOutline(id)!.outline as powerbi.visuals.GroupSubSelectionOutline).outlines;

            expect(outlines[0]).toMatchObject({ x: 15, y: 15, width: 40, height: 40 });
        });

        it("adds a clip path when the restricting element clips", () => {
            const helper = createHelper();
            const element = createRestrictedElement(SubSelectionOutlineRestrictionType.Clip);

            const id = helper.updateElementOutline(element, SubSelectionOutlineVisibility.Active);
            const outlines = (helper.getRegionOutline(id)!.outline as powerbi.visuals.GroupSubSelectionOutline).outlines;

            expect(outlines[0]).toMatchObject({
                x: 0,
                y: 0,
                width: 100,
                height: 100,
                clipPath: { type: SubSelectionOutlineType.Rectangle, x: 10, y: 10, width: 50, height: 50 },
            });
        });

        it("ignores restricting elements that belong to another helper", () => {
            const foreignHost = document.createElement("div");
            foreignHost.setAttribute(SubSelectableRestrictingElementAttribute, `${SubSelectionOutlineRestrictionType.Clamp}`);
            setRect(foreignHost, { x: 10, y: 10, width: 50, height: 50 });
            document.body.appendChild(foreignHost);
            foreignHost.appendChild(host);

            const helper = createHelper();
            const element = createSubSelectable(host, { objectName: "text", rect: { x: 0, y: 0, width: 100, height: 100 } });

            const id = helper.updateElementOutline(element, SubSelectionOutlineVisibility.Active);
            const outlines = (helper.getRegionOutline(id)!.outline as powerbi.visuals.GroupSubSelectionOutline).outlines;

            expect(outlines[0]).toMatchObject({ x: 0, y: 0, width: 100, height: 100 });
        });
    });

    describe("region outlines", () => {
        const outline = (id: string, visibility: number) => ({
            id,
            visibility,
            outline: { type: SubSelectionOutlineType.Rectangle, x: 0, y: 0, width: 1, height: 1 },
        }) as unknown as Parameters<HtmlSubSelectionHelper["updateRegionOutline"]>[0];

        it("stores and returns outlines by id", () => {
            const helper = createHelper();

            helper.updateRegionOutlines([outline("a", SubSelectionOutlineVisibility.Active), outline("b", SubSelectionOutlineVisibility.Hover)]);

            expect(helper.getRegionOutlines(["a", "b", "missing"] as powerbi.visuals.SubSelectionRegionOutlineId[]).map(o => o?.visibility)).toEqual([
                SubSelectionOutlineVisibility.Active,
                SubSelectionOutlineVisibility.Hover,
                undefined,
            ]);
        });

        it("returns a copy from getAllOutlines", () => {
            const helper = createHelper();
            helper.updateRegionOutline(outline("a", SubSelectionOutlineVisibility.Active));

            const all = helper.getAllOutlines();
            delete all["a"];

            expect(helper.getRegionOutline("a" as powerbi.visuals.SubSelectionRegionOutlineId)).toBeDefined();
        });

        it("hides every outline", () => {
            const helper = createHelper();
            helper.updateRegionOutlines([outline("a", SubSelectionOutlineVisibility.Active), outline("b", SubSelectionOutlineVisibility.Hover)]);

            helper.hideAllOutlines();

            expect(Object.values(helper.getAllOutlines()).map(o => o.visibility)).toEqual([
                SubSelectionOutlineVisibility.None,
                SubSelectionOutlineVisibility.None,
            ]);
        });

        it("clears only the hovered outline", () => {
            const helper = createHelper();
            helper.updateRegionOutlines([outline("a", SubSelectionOutlineVisibility.Active), outline("b", SubSelectionOutlineVisibility.Hover)]);

            helper.clearHoveredOutline();

            expect(helper.getRegionOutline("a" as powerbi.visuals.SubSelectionRegionOutlineId)!.visibility).toBe(SubSelectionOutlineVisibility.Active);
            expect(helper.getRegionOutline("b" as powerbi.visuals.SubSelectionRegionOutlineId)!.visibility).toBe(SubSelectionOutlineVisibility.None);
        });
    });

    describe("createVisualSubSelectionForSingleObject", () => {
        it("creates a sub-selection for the given object", () => {
            const helper = createHelper();
            const selectionId = createSelectionId("key1");

            const subSelection = helper.createVisualSubSelectionForSingleObject({
                objectName: "text",
                subSelectionType: SubSelectionStylesType.Shape,
                displayName: "Text",
                showUI: true,
                selectionId,
                selectionOrigin: { x: 5, y: 6 },
                focusOrder: 2,
                metadata: { custom: true },
            });

            expect(subSelection).toEqual({
                customVisualObjects: [{ objectName: "text", selectionId }],
                showUI: true,
                displayName: "Text",
                subSelectionType: SubSelectionStylesType.Shape,
                selectionOrigin: { x: 5, y: 6 },
                metadata: { custom: true },
                focusOrder: 2,
            });
        });

        it("omits optional members when they are not provided", () => {
            const helper = createHelper();

            const subSelection = helper.createVisualSubSelectionForSingleObject({
                objectName: "text",
                subSelectionType: SubSelectionStylesType.Shape,
                displayName: "Text",
                showUI: false,
            });

            expect(subSelection.customVisualObjects[0].selectionId).toBeUndefined();
            expect(subSelection.selectionOrigin).toBeUndefined();
            expect("metadata" in subSelection).toBe(false);
            expect("focusOrder" in subSelection).toBe(false);
        });

        it.each([
            ["Text", SubSelectionStylesType.Text],
            ["NumericText", SubSelectionStylesType.NumericText],
        ])("anchors the selection origin to the top of the text for %s", (_name, subSelectionType) => {
            const helper = createHelper();

            const subSelection = helper.createVisualSubSelectionForSingleObject({
                objectName: "text",
                subSelectionType,
                displayName: "Text",
                showUI: false,
                selectionOrigin: { x: 5, y: 6 },
            });

            expect(subSelection.selectionOrigin).toEqual({ x: 5, y: 6, offset: { x: 0, y: -6 } });
        });

        it("keeps the selection origin as is for non text types", () => {
            const helper = createHelper();

            const subSelection = helper.createVisualSubSelectionForSingleObject({
                objectName: "shape",
                subSelectionType: SubSelectionStylesType.Shape,
                displayName: "Shape",
                showUI: false,
                selectionOrigin: { x: 5, y: 6 },
            });

            expect(subSelection.selectionOrigin).toEqual({ x: 5, y: 6 });
        });
    });

    describe("format mode and events", () => {
        it("sub-selects the closest sub-selectable element on click", () => {
            const helper = createHelper();
            const element = createSubSelectable(host, {
                objectName: "text",
                displayName: "Text",
                subSelectionType: SubSelectionStylesType.Shape,
            });
            const child = document.createElement("span");
            element.appendChild(child);

            helper.setFormatMode(true);
            dispatch(child, "click", { clientX: 5, clientY: 7 });

            expect(subSelectionService.subSelect).toHaveBeenCalledWith({
                customVisualObjects: [{ objectName: "text", selectionId: undefined }],
                showUI: false,
                displayName: "Text",
                subSelectionType: SubSelectionStylesType.Shape,
                selectionOrigin: { x: 5, y: 7 },
            });
        });

        it("shows the UI when the context menu is used", () => {
            const helper = createHelper();
            const element = createSubSelectable(host, { objectName: "text", displayName: "Text", subSelectionType: SubSelectionStylesType.Shape });

            helper.setFormatMode(true);
            dispatch(element, "contextmenu", { clientX: 1, clientY: 2 });

            expect(subSelectionService.subSelect).toHaveBeenCalledWith(expect.objectContaining({ showUI: true }));
        });

        it("clears the sub-selection when clicking outside of a sub-selectable element", () => {
            const helper = createHelper();
            createSubSelectable(host, { objectName: "text" });

            helper.setFormatMode(true);
            dispatch(host, "click", { clientX: 3, clientY: 4 });

            expect(subSelectionService.subSelect).toHaveBeenCalledWith({
                customVisualObjects: [],
                displayName: "",
                showUI: false,
                subSelectionType: SubSelectionStylesType.Shape,
                selectionOrigin: { x: 3, y: 4 },
            });
        });

        it("passes the metadata from the metadata callback", () => {
            const subSelectionMetadataCallback = vi.fn().mockReturnValue({ index: 3 });
            const helper = createHelper({ subSelectionMetadataCallback });
            const element = createSubSelectable(host, { objectName: "text" });

            helper.setFormatMode(true);
            dispatch(element, "click");

            expect(subSelectionMetadataCallback).toHaveBeenCalledWith(element);
            expect(subSelectionService.subSelect).toHaveBeenCalledWith(expect.objectContaining({ metadata: { index: 3 } }));
        });

        it("does not react to events when the format mode is off", () => {
            const helper = createHelper();
            const element = createSubSelectable(host, { objectName: "text" });

            helper.setFormatMode(true);
            helper.setFormatMode(false);
            dispatch(element, "click");

            expect(subSelectionService.subSelect).not.toHaveBeenCalled();
        });

        it("ignores repeated calls with the same format mode", () => {
            const helper = createHelper();

            helper.setFormatMode(false);
            subSelectionService.updateRegionOutlines.mockClear();
            helper.setFormatMode(false);

            expect(subSelectionService.updateRegionOutlines).not.toHaveBeenCalled();
        });

        it("shows a hover outline on pointer over and removes it on pointer leave", () => {
            const helper = createHelper();
            const element = createSubSelectable(host, { objectName: "text", rect: { x: 0, y: 0, width: 10, height: 10 } });

            helper.setFormatMode(true);
            dispatch(element, "pointerover");
            expect(helper.getRegionOutline("text" as powerbi.visuals.SubSelectionRegionOutlineId)!.visibility).toBe(SubSelectionOutlineVisibility.Hover);

            element.dispatchEvent(new MouseEvent("pointerleave"));
            expect(helper.getRegionOutline("text" as powerbi.visuals.SubSelectionRegionOutlineId)!.visibility).toBe(SubSelectionOutlineVisibility.None);
        });

        it("does not create a hover outline for elements that hide their outline", () => {
            const helper = createHelper();
            const element = createSubSelectable(host, {
                objectName: "text",
                attributes: { [SubSelectableHideOutlineAttribute]: "true" },
            });

            helper.setFormatMode(true);
            dispatch(element, "pointerover");

            expect(helper.getRegionOutline("text" as powerbi.visuals.SubSelectionRegionOutlineId)).toBeUndefined();
        });

        it("removes the handlers and hides the outlines on destroy", () => {
            const helper = createHelper();
            const element = createSubSelectable(host, { objectName: "text" });
            helper.setFormatMode(true);
            helper.updateElementOutline(element, SubSelectionOutlineVisibility.Active);

            helper.destroy();
            dispatch(element, "click");

            expect(subSelectionService.subSelect).not.toHaveBeenCalled();
            expect(helper.getRegionOutline("text" as powerbi.visuals.SubSelectionRegionOutlineId)!.visibility).toBe(SubSelectionOutlineVisibility.None);
        });
    });

    describe("getSubSelectionSourceFromEvent", () => {
        it("returns the sub-selectable element and its sub-selection", () => {
            const helper = createHelper();
            const element = createSubSelectable(host, { objectName: "text", displayName: "Text" });

            const event = new MouseEvent("click", { bubbles: true });
            let source: ReturnType<HtmlSubSelectionHelper["getSubSelectionSourceFromEvent"]>;
            element.addEventListener("click", e => {
                source = helper.getSubSelectionSourceFromEvent(e as PointerEvent);
            });
            element.dispatchEvent(event);

            expect(source!.subSelectionElement).toBe(element);
            expect(source!.visualSubSelection.displayName).toBe("Text");
        });

        it("returns undefined when the event does not originate from a sub-selectable element", () => {
            const helper = createHelper();

            let source: ReturnType<HtmlSubSelectionHelper["getSubSelectionSourceFromEvent"]>;
            host.addEventListener("click", e => {
                source = helper.getSubSelectionSourceFromEvent(e as PointerEvent);
            });
            host.dispatchEvent(new MouseEvent("click", { bubbles: true }));

            expect(source!).toBeUndefined();
        });
    });

    describe("getAllSubSelectables", () => {
        it("returns unique sub-selections ordered by their vertical position", () => {
            const helper = createHelper();
            createSubSelectable(host, { objectName: "bottom", displayName: "Bottom", rect: { x: 0, y: 100, width: 10, height: 10 } });
            createSubSelectable(host, { objectName: "top", displayName: "Top", rect: { x: 0, y: 0, width: 10, height: 10 } });
            createSubSelectable(host, { objectName: "top", displayName: "Top", rect: { x: 20, y: 0, width: 10, height: 10 } });

            const subSelections = helper.getAllSubSelectables();

            expect(subSelections.map(subSelection => subSelection.displayName)).toEqual(["Top", "Bottom"]);
            expect(subSelections[0].selectionOrigin).toEqual({ x: 5, y: 5 });
            expect(subSelections[1].selectionOrigin).toEqual({ x: 5, y: 105 });
        });

        it("keeps elements with the same object name but different selection ids", () => {
            const first = createSubSelectable(host, { objectName: "text", rect: { x: 0, y: 0, width: 10, height: 10 } });
            const second = createSubSelectable(host, { objectName: "text", rect: { x: 0, y: 20, width: 10, height: 10 } });
            const selectionIds = new Map([[first, createSelectionId("a")], [second, createSelectionId("b")]]);
            const helper = createHelper({ selectionIdCallback: element => selectionIds.get(element)! });

            expect(helper.getAllSubSelectables()).toHaveLength(2);
        });

        it("filters the sub-selectables by type", () => {
            const helper = createHelper();
            createSubSelectable(host, { objectName: "text", subSelectionType: SubSelectionStylesType.Text });
            createSubSelectable(host, { objectName: "shape", subSelectionType: SubSelectionStylesType.Shape });

            const subSelections = helper.getAllSubSelectables(SubSelectionStylesType.Shape);

            expect(subSelections).toHaveLength(1);
            expect(subSelections[0].customVisualObjects[0].objectName).toBe("shape");
        });

        it("filters sub-selectables by the zero-valued None type", () => {
            const helper = createHelper();
            createSubSelectable(host, { objectName: "none", subSelectionType: SubSelectionStylesType.None });
            createSubSelectable(host, { objectName: "shape", subSelectionType: SubSelectionStylesType.Shape });
            createSubSelectable(host, { objectName: "text", subSelectionType: SubSelectionStylesType.Text });

            const subSelections = helper.getAllSubSelectables(SubSelectionStylesType.None);

            expect(subSelections).toHaveLength(1);
            expect(subSelections[0].customVisualObjects[0].objectName).toBe("none");
        });

        it("defaults the sub-selection type to None when the element has no type attribute", () => {
            const helper = createHelper();
            createSubSelectable(host, { objectName: "untyped", rect: { x: 0, y: 0, width: 10, height: 10 } });

            const subSelections = helper.getAllSubSelectables();

            expect(subSelections).toHaveLength(1);
            expect(subSelections[0].subSelectionType).toBe(SubSelectionStylesType.None);
        });

        it("defaults the sub-selection type to None when the type attribute is not a known value", () => {
            const helper = createHelper();
            const element = createSubSelectable(host, { objectName: "broken", rect: { x: 0, y: 0, width: 10, height: 10 } });
            element.setAttribute("data-sub-selection-type", "shape");

            const subSelections = helper.getAllSubSelectables();

            expect(subSelections).toHaveLength(1);
            expect(subSelections[0].subSelectionType).toBe(SubSelectionStylesType.None);
        });

        it("ignores sub-selectable elements owned by another helper", () => {
            const helper = createHelper();
            const foreignHost = document.createElement("div");
            foreignHost.setAttribute("data-helper-host", "true");
            host.appendChild(foreignHost);
            const foreignElement = document.createElement("div");
            foreignElement.classList.add(HtmlSubSelectableClass);
            foreignElement.setAttribute(SubSelectableObjectNameAttribute, "foreign");
            setRect(foreignElement, { width: 10, height: 10 });
            foreignHost.appendChild(foreignElement);

            expect(helper.getAllSubSelectables()).toEqual([]);
        });
    });

    describe("sub-selection state", () => {
        it("marks the matching elements as sub-selected", () => {
            const helper = createHelper();
            const matching = createSubSelectable(host, { objectName: "text" });
            const other = createSubSelectable(host, { objectName: "shape" });

            helper.setSubSelectedStateDOM([{ customVisualObjects: [{ objectName: "text" }] } as powerbi.visuals.CustomVisualSubSelection]);

            expect(matching.getAttribute(SubSelectableSubSelectedAttribute)).toBe("true");
            expect(other.hasAttribute(SubSelectableSubSelectedAttribute)).toBe(false);
        });

        it("matches elements through their alternative object name", () => {
            const helper = createHelper();
            const element = createSubSelectable(host, {
                objectName: "markers",
                attributes: { [SubSelectableAltObjectNameAttribute]: "lines" },
            });

            const elements = helper.getElementsFromSubSelections([
                { customVisualObjects: [{ objectName: "lines" }] } as powerbi.visuals.CustomVisualSubSelection,
            ]);

            expect(elements).toEqual([element]);
        });

        it("uses the selection id callback to disambiguate elements", () => {
            const first = createSubSelectable(host, { objectName: "text" });
            const second = createSubSelectable(host, { objectName: "text" });
            const selectionIds = new Map([[first, createSelectionId("a")], [second, createSelectionId("b")]]);
            const helper = createHelper({ selectionIdCallback: element => selectionIds.get(element)! });

            const elements = helper.getElementsFromSubSelections([
                { customVisualObjects: [{ objectName: "text", selectionId: createSelectionId("b") }] } as powerbi.visuals.CustomVisualSubSelection,
            ]);

            expect(elements).toEqual([second]);
        });

        it("returns an empty array when there are no sub-selections", () => {
            const helper = createHelper();
            createSubSelectable(host, { objectName: "text" });

            expect(helper.getElementsFromSubSelections(undefined as unknown as powerbi.visuals.CustomVisualSubSelection[])).toEqual([]);
        });
    });

    describe("updateOutlinesFromSubSelections", () => {
        const subSelection = { customVisualObjects: [{ objectName: "text" }] } as powerbi.visuals.CustomVisualSubSelection;

        it("creates active outlines for the sub-selected elements", () => {
            const helper = createHelper();
            createSubSelectable(host, { objectName: "text", rect: { x: 0, y: 0, width: 10, height: 10 } });

            helper.updateOutlinesFromSubSelections([subSelection]);

            expect(helper.getRegionOutline("text" as powerbi.visuals.SubSelectionRegionOutlineId)!.visibility).toBe(SubSelectionOutlineVisibility.Active);
        });

        it("skips elements that hide their outline", () => {
            const helper = createHelper();
            createSubSelectable(host, { objectName: "text", attributes: { [SubSelectableHideOutlineAttribute]: "true" } });

            helper.updateOutlinesFromSubSelections([subSelection]);

            expect(helper.getRegionOutline("text" as powerbi.visuals.SubSelectionRegionOutlineId)).toBeUndefined();
        });

        it("clears the existing outlines when requested", () => {
            const helper = createHelper();
            const element = createSubSelectable(host, { objectName: "shape", rect: { x: 0, y: 0, width: 10, height: 10 } });
            helper.updateElementOutline(element, SubSelectionOutlineVisibility.Active);

            helper.updateOutlinesFromSubSelections([subSelection], true /* clearExistingOutlines */);

            expect(helper.getRegionOutline("shape" as powerbi.visuals.SubSelectionRegionOutlineId)!.visibility).toBe(SubSelectionOutlineVisibility.None);
        });

        it("does not render when the render is suppressed", () => {
            const helper = createHelper();
            createSubSelectable(host, { objectName: "text" });

            helper.updateOutlinesFromSubSelections([subSelection], false /* clearExistingOutlines */, true /* suppressRender */);

            expect(subSelectionService.updateRegionOutlines).not.toHaveBeenCalled();
        });

        it("recomputes the outlines of the last sub-selections on refresh", () => {
            const helper = createHelper();
            const element = createSubSelectable(host, { objectName: "text", rect: { x: 0, y: 0, width: 10, height: 10 } });
            helper.updateOutlinesFromSubSelections([subSelection]);

            setRect(element, { x: 0, y: 0, width: 50, height: 50 });
            helper.refreshOutlines();

            const outlines = (helper.getRegionOutline("text" as powerbi.visuals.SubSelectionRegionOutlineId)!.outline as powerbi.visuals.GroupSubSelectionOutline).outlines;
            expect(outlines[0]).toMatchObject({ width: 50, height: 50 });
        });
    });

    describe("custom outlines", () => {
        const subSelection = { customVisualObjects: [{ objectName: "text" }] } as powerbi.visuals.CustomVisualSubSelection;
        const customOutline = {
            id: "custom",
            type: SubSelectionOutlineType.Rectangle,
            x: 0,
            y: 0,
            width: 10,
            height: 10,
        } as unknown as powerbi.visuals.SubSelectionRegionOutlineFragment;

        it("uses the custom outline callback instead of the element outlines", () => {
            const helper = createHelper({ customOutlineCallback: () => [customOutline] });
            createSubSelectable(host, { objectName: "text", rect: { x: 0, y: 0, width: 10, height: 10 } });

            helper.updateOutlinesFromSubSelections([subSelection]);

            expect(helper.getRegionOutline("custom" as powerbi.visuals.SubSelectionRegionOutlineId)!.visibility).toBe(SubSelectionOutlineVisibility.Active);
            expect(helper.getRegionOutline("text" as powerbi.visuals.SubSelectionRegionOutlineId)).toBeUndefined();
        });

        it("returns the sub-selections that produced custom outlines", () => {
            const helper = createHelper({ customOutlineCallback: () => [customOutline] });

            expect(helper.updateCustomOutlinesFromSubSelections([subSelection])).toEqual([subSelection]);
        });

        it("returns nothing when the callback produces no outlines", () => {
            const helper = createHelper({ customOutlineCallback: () => [] });

            expect(helper.updateCustomOutlinesFromSubSelections([subSelection])).toEqual([]);
        });

        it("keeps an active custom outline active while hovering", () => {
            const helper = createHelper({ customOutlineCallback: () => [customOutline] });
            helper.updateCustomOutlinesFromSubSelections([subSelection], SubSelectionOutlineVisibility.Active);

            helper.updateCustomOutlinesFromSubSelections([subSelection], SubSelectionOutlineVisibility.Hover);

            expect(helper.getRegionOutline("custom" as powerbi.visuals.SubSelectionRegionOutlineId)!.visibility).toBe(SubSelectionOutlineVisibility.Active);
        });

        it("can be registered after the helper was created", () => {
            const helper = createHelper();

            helper.setCustomOutlineCallback(() => [customOutline]);

            expect(helper.updateCustomOutlinesFromSubSelections([subSelection])).toEqual([subSelection]);
        });
    });

    describe("custom element callback", () => {
        it("outlines the elements returned by the callback", () => {
            const custom = createSubSelectable(host, { objectName: "custom", rect: { x: 0, y: 0, width: 10, height: 10 } });
            const helper = createHelper({ customElementCallback: () => [custom] });
            const element = createSubSelectable(host, { objectName: "text", rect: { x: 0, y: 0, width: 10, height: 10 } });

            helper.setFormatMode(true);
            dispatch(element, "pointerover");

            expect(helper.getRegionOutline("custom" as powerbi.visuals.SubSelectionRegionOutlineId)!.visibility).toBe(SubSelectionOutlineVisibility.Hover);
        });
    });

    describe("onVisualScroll", () => {
        it("clears the sub-selection while scrolling and restores it afterwards", () => {
            vi.useFakeTimers();
            const helper = createHelper();
            createSubSelectable(host, { objectName: "text", rect: { x: 0, y: 0, width: 10, height: 10 } });
            const subSelection = { customVisualObjects: [{ objectName: "text" }] } as powerbi.visuals.CustomVisualSubSelection;
            helper.updateOutlinesFromSubSelections([subSelection]);

            helper.onVisualScroll();
            expect(subSelectionService.subSelect).toHaveBeenCalledWith(undefined);

            helper.onVisualScroll();
            vi.advanceTimersByTime(100);

            expect(subSelectionService.subSelect).toHaveBeenCalledTimes(2);
            expect(subSelectionService.subSelect).toHaveBeenLastCalledWith(subSelection);
        });

        it("does not restore anything when there is no sub-selection", () => {
            vi.useFakeTimers();
            const helper = createHelper();

            helper.onVisualScroll();
            vi.advanceTimersByTime(100);

            expect(subSelectionService.subSelect).toHaveBeenCalledTimes(1);
            expect(subSelectionService.subSelect).toHaveBeenCalledWith(undefined);
        });
    });
});

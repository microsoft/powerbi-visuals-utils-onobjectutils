import { select, selectAll, Selection } from "d3-selection";
import powerbi from "powerbi-visuals-api";

import { CreateVisualSubSelectionFromObjectArgs, HtmlSubselectionHelperArgs, HtmlSubSelectionSource, ISubSelectionHelper, SubSelectionElementData } from "./types";
import { debounce, getObjectValues, groupArrayElements, isArrayEmpty, isEqual, getUniques, equalsSelectionId } from "./helperFunctions";

import CustomVisualSubSelection = powerbi.visuals.CustomVisualSubSelection;
import GroupSubSelectionOutline = powerbi.visuals.GroupSubSelectionOutline;
import HelperSubSelectionRegionOutline = powerbi.visuals.HelperSubSelectionRegionOutline
import IPoint = powerbi.extensibility.IPoint;
import ISelectionId = powerbi.visuals.ISelectionId
import IVisualSubSelectionService = powerbi.extensibility.IVisualSubSelectionService
import RectangleSubSelectionOutline = powerbi.visuals.RectangleSubSelectionOutline;
import SubSelectionOutlineRestrictionType = powerbi.visuals.SubSelectionOutlineRestrictionType;
import SubSelectionOutlineType = powerbi.visuals.SubSelectionOutlineType;
import SubSelectionOutlineVisibility = powerbi.visuals.SubSelectionOutlineVisibility;
import SubSelectionRegionOutlineFragment = powerbi.visuals.SubSelectionRegionOutlineFragment;
import SubSelectionRegionOutlineId = powerbi.visuals.SubSelectionRegionOutlineId;
import SubSelectionStylesType = powerbi.visuals.SubSelectionStylesType;

const d3 = { select, selectAll };

// Used on the sub-selectable element
const eventSuffix = 'htmlSubSelection';
const subSelectableClassAndSelector = createClassAndSelector('sub-selectable');
const SubSelectionData = 'sub-selection-data';
export const DirectEditPlaceholderClassAndSelector = createClassAndSelector('direct-edit-placeholder');
export const DirectEditPlaceholderOutlineClassAndSelector = createClassAndSelector('direct-edit-placeholder-outline');
export const FormatModeAttribute = 'format-mode';
export const HtmlSubSelectableClass = subSelectableClassAndSelector.class;
export const HtmlSubSelectableSelector = subSelectableClassAndSelector.selector;
export const SubSelectableDisplayNameAttribute = 'data-sub-selection-display-name';
export const SubSelectableHideOutlineAttribute = 'data-sub-selection-hide-outline';
export const SubSelectableObjectNameAttribute = 'data-sub-selection-object-name';
// Used for when another object is associated with the sub-selectable element (e.g. Lines/Markers on interactivity lines)
export const SubSelectableAltObjectNameAttribute = 'data-sub-selection-alt-object-name';
export const SubSelectableTypeAttribute = 'data-sub-selection-type';
export const SubSelectableDirectEdit = 'data-sub-selection-direct-edit';
export const SubSelectableSubSelectedAttribute = 'data-sub-selection-sub-selected';
export const SubSelectableUIAnchorAttribute = 'data-sub-selection-ui-anchor';
// Used to indicate the element which will restricted the outlines and the type of restriction (clamp or clip)
export const SubSelectableRestrictingElementAttribute = 'data-sub-selection-restricting-element';
/** Indicates that this helper has attached to this element */
const helperHostAttribute = 'data-helper-host';
const scrollDebounceInterval = 100;

interface ClassAndSelector {
    class: string;
    selector: string;
}

interface IRawRect<T> {
    left: T;
    top: T;
    width: T;
    height: T;
}

type IRect = IRawRect<number>;

function getEventName(eventName: string): string {
    return `${eventName}.${eventSuffix}`;
}

function helperOwnsElement(host: HTMLElement, element: HTMLElement): boolean {
    return element.closest(`[${helperHostAttribute}]`) === host;
}

function createClassAndSelector(className: string): ClassAndSelector {
    return {
        class: className,
        selector: '.' + className,
    };
}

export class HtmlSubSelectionHelper implements ISubSelectionHelper<HTMLElement, CustomVisualSubSelection> {
    private readonly host: Selection<HTMLElement, unknown, null, unknown>;

    private customOutlineCallback: ((subSelection: CustomVisualSubSelection) => SubSelectionRegionOutlineFragment[]) | undefined;
    private formatMode: boolean;
    private readonly customElementCallback: ((subSelection: CustomVisualSubSelection) => HTMLElement[]) | undefined;
    private readonly hostElement: HTMLElement;
    private readonly selectionIdCallback: ((e: HTMLElement) => ISelectionId) | undefined;
    private readonly subSelectionMetadataCallback: ((element: HTMLElement) => unknown) | undefined;
    private readonly subSelectionService: IVisualSubSelectionService;
    private subSelectionRegionOutlines: Record<string, HelperSubSelectionRegionOutline>;
    // maintain some state for reapplciation of subselections (e.g. scroll)
    private subSelections: CustomVisualSubSelection[];
    private scrollSubSelections: CustomVisualSubSelection[];
    private scrollDebounce: () => void;

    private constructor(
        args: HtmlSubselectionHelperArgs,
    ) {
        this.hostElement = args.hostElement;
        this.host = d3.select(args.hostElement);
        this.subSelectionService = args.subSelectionService;
        this.subSelectionRegionOutlines = {};
        this.selectionIdCallback = args.selectionIdCallback;
        this.customOutlineCallback = args.customOutlineCallback;
        this.customElementCallback = args.customElementCallback;
        this.subSelectionMetadataCallback = args.subSelectionMetadataCallback;
        this.host.attr(helperHostAttribute, true);
    }

    public static createHtmlSubselectionHelper(args: HtmlSubselectionHelperArgs): HtmlSubSelectionHelper {
        return new HtmlSubSelectionHelper(args);
    }

    public setCustomOutlineCallback(customOutlineCallback: (subSelection: CustomVisualSubSelection) => SubSelectionRegionOutlineFragment[]): void {
        this.customOutlineCallback = customOutlineCallback;
    }

    public destroy(): void {
        this.removeEventHandlers();
        this.hideAllOutlines();
    }

    public setFormatMode(isFormatMode: boolean): void {
        if (this.formatMode === isFormatMode) {
            return;
        }

        this.formatMode = isFormatMode;
        if (isFormatMode) {
            this.attachEventHandlers();
        } else {
            this.removeEventHandlers();
            this.hideAllOutlines();
        }
    }

    private attachEventHandlers(): void {
        this.host.on(getEventName('pointerover'), (event) => {
            this.onPointerOver(event);
        });
        this.host.on(getEventName('click'), (event) => {
            this.onClick(event);
        });
        this.host.on(getEventName('contextmenu'), (event) => {
            this.onContextMenu(event);
        });
    }

    private removeEventHandlers(): void {
        this.host.on(getEventName('pointerover'), null);
        this.host.on(getEventName('click'), null);
        this.host.on(getEventName('contextmenu'), null);
    }

    public onVisualScroll(): void {
        // if scrolling
        if (!this.scrollDebounce) {
            this.onVisualScrollStart();
            this.scrollDebounce = debounce(() => this.onVisualScrollEnd(), scrollDebounceInterval);
        }

        this.scrollDebounce();
    }

    // clear subselections and set up state
    private onVisualScrollStart(): void {
        if (this.scrollSubSelections) {
            return;
        }
        this.subSelectionService.subSelect(undefined);
        this.scrollSubSelections = this.subSelections;
    }

    // reapply subselections
    private onVisualScrollEnd(): void {
        if (this.scrollSubSelections && this.scrollSubSelections.length > 0) {
            this.subSelectionService.subSelect(this.scrollSubSelections[0]);
        }
        this.scrollSubSelections = undefined;
        this.scrollDebounce = undefined;
    }

    private onPointerOver(event: PointerEvent): void {
        this.clearHoveredOutline();
        const subSelectionSource =
            this.getSubSelectionSourceFromEvent(event);
        if (subSelectionSource && subSelectionSource.subSelectionElement) {
            // If visualSubSelection has custom outlines, omit default behavior
            // If custom outline is already active, don't set to hover
            const visualSubSelectionsWithCustomOutlines = this.updateCustomOutlinesFromSubSelections([subSelectionSource.visualSubSelection], SubSelectionOutlineVisibility.Hover);
            if (visualSubSelectionsWithCustomOutlines.length === 0) {
                let elementsToUpdate = this.getSubSelectionElementsFromSubSelectionSource(subSelectionSource);

                // Skip sub-selected elements or elements which should not be outlined
                elementsToUpdate = elementsToUpdate.filter(e => {
                    const selectedElement = d3.select(e);
                    return !selectedElement.attr(SubSelectableSubSelectedAttribute) && selectedElement.attr(SubSelectableHideOutlineAttribute) !== "true";
                });
                this.updateOutlinesFromSubSelectionElements(elementsToUpdate, SubSelectionOutlineVisibility.Hover);
            }

            this.renderOutlines();
            const pointerLeaveEventName = getEventName('pointerleave');
            const targetedElement = d3.select(subSelectionSource.subSelectionElement);
            // Attach a listener for leaving the sub-selected element
            // For entry, we care about hovering over any of the children of sub-selectable elements, so we can listen to all events and filter
            // For exit, we only want to react when you move outside of the sub-selection (so far).
            // That's more difficult with a single top-level handler and not storing state, so going with attaching events for now
            targetedElement.on(pointerLeaveEventName, () => {
                // Skip sub-selected elements
                if (targetedElement.attr(SubSelectableSubSelectedAttribute)) {
                    return;
                }

                this.clearHoveredOutline();

                targetedElement.on(pointerLeaveEventName, null);
            });
        }
    }

    public clearHoveredOutline(): void {
        const regionOutlines = getObjectValues(this.subSelectionRegionOutlines);
        const hoveredOutline = regionOutlines.find(outline => outline.visibility === SubSelectionOutlineVisibility.Hover);
        if (hoveredOutline) {
            this.subSelectionRegionOutlines[hoveredOutline.id] = {
                ...this.subSelectionRegionOutlines[hoveredOutline.id],
                visibility: SubSelectionOutlineVisibility.None
            };
            this.renderOutlines();
        }
    }

    private onClick(event: PointerEvent): void {
        this.subSelectFromEvent(event, false /* showUI */);
    }

    private onContextMenu(event: PointerEvent): void {
        this.subSelectFromEvent(event, true /* showUI */);
    }

    private subSelectFromEvent(event: PointerEvent, showUI: boolean): void {
        event.preventDefault();
        const newSubSelectionElements = this.getSubSelectionElementsFromEvent(event);
        // Mark the event as handled so containers don't process this event
        const selectionOrigin = {
            x: event.clientX,
            y: event.clientY,
        };
        if (isArrayEmpty(newSubSelectionElements)) {
            this.subSelectionService.subSelect({
                customVisualObjects: [],
                selectionOrigin,
                showUI
            });
            return;
        }

        const newSubSelection = newSubSelectionElements[0];
        const args = this.getCreateVisualSubSelectionArgs(event, newSubSelection, showUI)
        const subSelection = this.createVisualSubSelectionForSingleObject(args);
        this.scrollSubSelections = undefined;
        this.subSelectionService.subSelect(subSelection);
    }

    private getSubSelectionElementsFromEvent(event: PointerEvent): HTMLElement[] {
        const subSelectionSource = this.getSubSelectionSourceFromEvent(event);
        const subSelectionElements = this.getSubSelectionElementsFromSubSelectionSource(subSelectionSource);
        return subSelectionElements;
    }

    public getSubSelectionSourceFromEvent(event: PointerEvent): HtmlSubSelectionSource | undefined {
        const fullPath = event.composedPath();
        if (!fullPath) {
            return undefined;
        }

        // Find the root element in the path, remove everything above it
        const eventHandlerElementIndex = fullPath.indexOf(this.hostElement);
        const path = fullPath.slice(0, eventHandlerElementIndex + 1);
        let subSelectionElement: HTMLElement | undefined;
        // Use the closest parent to the event
        for (const currentElement of path as HTMLElement[]) {
            const currentSelection = d3.select(currentElement);

            // Only supports one level for now
            if (currentSelection.classed(HtmlSubSelectableClass)) {
                subSelectionElement = currentElement;
                break;
            }
        }

        if (subSelectionElement && helperOwnsElement(this.hostElement, subSelectionElement)) {
            const args = this.getCreateVisualSubSelectionArgs(event, subSelectionElement, false /**showUI */)
            const visualSubSelection = this.createVisualSubSelectionForSingleObject(args);

            return { subSelectionElement, visualSubSelection };
        }
        return undefined;
    }

    private getCreateVisualSubSelectionArgs(event: PointerEvent, subSelectionElement: HTMLElement, showUI: boolean): CreateVisualSubSelectionFromObjectArgs {
        const selectionId = this.selectionIdCallback ? this.selectionIdCallback(subSelectionElement) : undefined;
        const objectName = d3.select(subSelectionElement).attr(SubSelectableObjectNameAttribute);
        const displayName = this.getDisplayNameFromElement(subSelectionElement);
        const subSelectionType = this.getSubSelectionTypeFromElement(subSelectionElement);
        const selectionOrigin = {
            x: event.clientX,
            y: event.clientY,
        };
        const metadata = this.subSelectionMetadataCallback ? this.subSelectionMetadataCallback(subSelectionElement) : null;

        return {
            objectName,
            subSelectionType,
            displayName,
            showUI,
            selectionId,
            selectionOrigin,
            metadata,
        };
    }

    private getSubSelectionElementsFromSubSelectionSource(subSelectionSource: HtmlSubSelectionSource): HTMLElement[] {
        if (!subSelectionSource) {
            return [];
        }

        const { visualSubSelection } = subSelectionSource;
        if (this.customElementCallback) {
            const customElements = this.customElementCallback(visualSubSelection);
            if (!isArrayEmpty(customElements)) {
                return customElements;
            }
        }

        const subSelectables = this.getSubSelectableElements();
        const { objectName, selectionId } = visualSubSelection.customVisualObjects[0];
        let filteredSelectionElements = subSelectables.filter((subSelectable) => subSelectable.getAttribute(SubSelectableObjectNameAttribute) === objectName);
        if (this.selectionIdCallback) {
            const callback = (e: HTMLElement): ISelectionId => this.selectionIdCallback(e);
            filteredSelectionElements = filteredSelectionElements.filter((element) => equalsSelectionId(selectionId, callback(element)));
        }

        return filteredSelectionElements;
    }

    public updateElementOutline(element: HTMLElement, visibility: SubSelectionOutlineVisibility, suppressRender: boolean = false): SubSelectionRegionOutlineId {
        return this.updateElementOutlines([element], visibility, suppressRender)[0];
    }

    public updateElementOutlines(elements: HTMLElement[], visibility: SubSelectionOutlineVisibility, suppressRender: boolean = false,): SubSelectionRegionOutlineId[] {
        // Group up the elements into their region
        const elementsByOutlineRegionId = groupArrayElements(elements, element => {
            const subSelectedElement = d3.select(element);
            const regionId = this.getElementRegionOutlineId(subSelectedElement);
            return regionId;
        });
        const regionOutlineIds = Object.keys(elementsByOutlineRegionId) as SubSelectionRegionOutlineId[];
        for (const regionOutlineId of regionOutlineIds) {
            const subSelectionRegionOutline = this.getSubSelectionRegionOutline(
                regionOutlineId,
                elementsByOutlineRegionId[regionOutlineId],
                visibility
            );
            this.subSelectionRegionOutlines[regionOutlineId] = subSelectionRegionOutline;
        }
        if (!suppressRender) {
            this.renderOutlines();
        }

        return regionOutlineIds;
    }

    private getSubSelectionRegionOutline(id: SubSelectionRegionOutlineId, elements: HTMLElement[], visibility: SubSelectionOutlineVisibility): HelperSubSelectionRegionOutline {
        const outlines: RectangleSubSelectionOutline[] = [];
        let regionClipElement: HTMLElement;
        let regionClampElement: HTMLElement;

        for (const element of elements) {
            let outline: RectangleSubSelectionOutline = this.getRectangleSubSelectionOutline(element);
            const currentClampRestriction = this.getRestrictionElement(element, SubSelectionOutlineRestrictionType.Clamp);
            regionClampElement = currentClampRestriction;
            if (regionClampElement) {
                const clampRect = this.applyElementRestriction(regionClampElement, element);
                const newX = Math.max(outline.x, clampRect.left);
                const newY = Math.max(outline.y, clampRect.top);
                outline = {
                    ...outline,
                    x: newX,
                    y: newY,
                    height: Math.min(clampRect.top + clampRect.height - newY, outline.height),
                    width: Math.min(clampRect.left + clampRect.width - newX, outline.width),
                };
            }

            const currentClipRestriction = this.getRestrictionElement(element, SubSelectionOutlineRestrictionType.Clip);
            regionClipElement = currentClipRestriction;
            if (regionClipElement) {
                const clipRect = this.applyElementRestriction(regionClipElement, element);
                outline = {
                    ...outline,
                    clipPath: {
                        type: SubSelectionOutlineType.Rectangle,
                        x: clipRect.left,
                        y: clipRect.top,
                        height: clipRect.height,
                        width: clipRect.width,
                    },
                };
            }

            if (outline.height > 0 && outline.width > 0) {
                outlines.push(outline);
            }
        }
        const groupOutline: GroupSubSelectionOutline = {
            type: SubSelectionOutlineType.Group,
            outlines,
        };
        return {
            id,
            visibility,
            outline: groupOutline,
        };
    }

    private getRestrictionElement(element: HTMLElement, type: SubSelectionOutlineRestrictionType): HTMLElement | undefined {
        const restrictionElement: HTMLElement = element.closest(`[${SubSelectableRestrictingElementAttribute}="${type}"]`);
        if (restrictionElement && helperOwnsElement(this.hostElement, restrictionElement)) {
            return restrictionElement;
        }
        return undefined;
    }

    private applyElementRestriction(restrictingElement: HTMLElement, subselectionElement: HTMLElement): IRect {
        const elementRect = restrictingElement.getBoundingClientRect();
        const rect: IRect = {
            top: elementRect.top,
            left: elementRect.left,
            height: elementRect.height,
            width: elementRect.width,
        };
        const data = HtmlSubSelectionHelper.getDataForElement(subselectionElement);
        if (data && data.outlineRestrictionOptions) {
            const { margin, padding } = data.outlineRestrictionOptions;
            let x = 0, y = 0, height = 0, width = 0;
            if (margin) {
                const { top, left, right, bottom } = margin;
                x += left;
                y += top;
                height -= (bottom + top);
                width -= (left + right);
            }

            if (padding) {
                const { top, left, right, bottom } = padding;
                x -= left;
                y -= top;
                height += (bottom + top);
                width += (left + right);
            }

            rect.left += x;
            rect.top += y;
            rect.height += height;
            rect.width += width;
        }
        return rect;
    }

    private getRectangleSubSelectionOutline(element: HTMLElement): RectangleSubSelectionOutline {
        const domRect = element.getBoundingClientRect();
        const { x, y, width, height } = domRect;
        const outline: RectangleSubSelectionOutline = {
            height,
            width,
            x,
            y,
            type: SubSelectionOutlineType.Rectangle,
        };
        if (element.hasAttribute(SubSelectableDirectEdit)) {
            outline.cVDirectEdit = JSON.parse(element.getAttribute(SubSelectableDirectEdit)!);
        }

        return outline;
    }

    public updateRegionOutline(regionOutline: HelperSubSelectionRegionOutline, suppressRender: boolean = false): void {
        this.updateRegionOutlines([regionOutline], suppressRender);
    }

    public updateRegionOutlines(regionOutlines: HelperSubSelectionRegionOutline[], suppressRender: boolean = false): void {
        for (const regionOutline of regionOutlines) {
            this.subSelectionRegionOutlines[regionOutline.id] = regionOutline;
        }
        if (!suppressRender) {
            this.renderOutlines();
        }
    }

    public getElementsFromSubSelections(subSelections: CustomVisualSubSelection[]): HTMLElement[] {
        if (!subSelections) {
            return [];
        }

        const elements: HTMLElement[] = [];
        // Update the sub-selection status on the elements
        const subSelectables = this.getSubSelectableElements();
        const isElementSubSelected = this.isElementSubSelected;
        const selectionIdCallbackFn = this.selectionIdCallback;
        d3.selectAll(subSelectables).each(
            function () {
                let isSubSelected = false;

                const element = this;

                isSubSelected = isElementSubSelected(element, subSelections, selectionIdCallbackFn!);
                if (isSubSelected) {
                    elements.push(element);
                }
            }
        );
        return elements;
    }

    // Updates the subselected attribute status on the elements associated with the VisualSubSelection
    public setSubSelectedStateDOM(subSelections: CustomVisualSubSelection[]): void {
        if (!subSelections) {
            return;
        }

        // Update the sub-selection status on the elements
        const subSelectables = this.getSubSelectableElements();
        const isElementSubSelected = this.isElementSubSelected;
        const selectionIdCallbackFn = this.selectionIdCallback;
        d3.selectAll(subSelectables).attr(
            SubSelectableSubSelectedAttribute,
            function () {
                let isSubSelected = false;

                const element = this;
                isSubSelected = isElementSubSelected(element, subSelections, selectionIdCallbackFn!);

                if (isSubSelected) {
                    return true;
                }
                return null;
            }
        );
    }

    private isElementSubSelected(element: HTMLElement, subSelections: CustomVisualSubSelection[], selectionIdCallbackFn: (e: HTMLElement) => ISelectionId): boolean {
        if (!subSelections) {
            return false;
        }

        const selectionObjectName = d3.select(element).attr(SubSelectableObjectNameAttribute);
        const selectionAltObjectName = d3.select(element).attr(SubSelectableAltObjectNameAttribute);
        const isSubSelected = subSelections.some(subSelection => subSelection.customVisualObjects?.some(customVisualObject => {
            let selectorMatches = true;
            if (selectionIdCallbackFn && customVisualObject.selectionId) {
                let otherSelectionId = selectionIdCallbackFn(element);
                if (!equalsSelectionId(customVisualObject.selectionId, otherSelectionId)) {
                    selectorMatches = false;
                }
            }

            return (customVisualObject.objectName === selectionObjectName || customVisualObject.objectName === selectionAltObjectName) && selectorMatches;
        }));

        return isSubSelected;
    }

    /**
     * @returns visualSubSelections that matched a custom outline
     */
    public updateCustomOutlinesFromSubSelections(subSelections: CustomVisualSubSelection[], visibility: SubSelectionOutlineVisibility = SubSelectionOutlineVisibility.Active): CustomVisualSubSelection[] {
        const visualSubSelectionsWithCustomOutlines: CustomVisualSubSelection[] = [];
        if (this.customOutlineCallback && !isArrayEmpty(subSelections)) {
            for (const subSelection of subSelections) {
                const customOutlines = this.customOutlineCallback(subSelection);
                if (customOutlines && customOutlines.length > 0) {
                    for (const outline of customOutlines) {
                        const currentOutlineVisibility = this.subSelectionRegionOutlines[outline.id]?.visibility;
                        // If the custom outline is recorded & already active, don't set it to hover, reaching 2nd clause implies visibility === SubSelectionOutlineVisibility.Hover
                        if (visibility !== SubSelectionOutlineVisibility.Hover || currentOutlineVisibility !== SubSelectionOutlineVisibility.Active) {
                            this.setOutline(outline, visibility);
                        }
                    }
                    visualSubSelectionsWithCustomOutlines.push(subSelection);
                }
            }
        }
        return visualSubSelectionsWithCustomOutlines;
    }

    public updateOutlinesFromSubSelectionElements(elementsToUpdate: HTMLElement[], visibility: SubSelectionOutlineVisibility = SubSelectionOutlineVisibility.Active): void {
        if (!isArrayEmpty(elementsToUpdate)) {
            this.updateElementOutlines(elementsToUpdate, visibility, true /* suppressRender */);
        }
    }

    public updateOutlinesFromSubSelections(subSelections: CustomVisualSubSelection[], clearExistingOutlines?: boolean, suppressRender?: boolean): void {
        if (clearExistingOutlines) {
            this.hideAllOutlines(true);
        }

        this.subSelections = subSelections;
        let subSelectionsNoCustomOutlines = subSelections ? [...subSelections] : [];
        // Set subselected state for both custom and regular flows together
        this.setSubSelectedStateDOM(subSelections);
        // If visualSubSelection has custom outlines, omit default behavior
        const visualSubSelectionsWithCustomOutlines = this.updateCustomOutlinesFromSubSelections(subSelections);
        if (visualSubSelectionsWithCustomOutlines?.length > 0) {
            subSelectionsNoCustomOutlines = subSelectionsNoCustomOutlines.filter((visualSubSelection) => !visualSubSelectionsWithCustomOutlines.find((vs) => isEqual(visualSubSelection, vs)));
        }

        const elements = this.getElementsFromSubSelections(subSelectionsNoCustomOutlines);
        const elementsToShow = elements.filter((e) => d3.select(e).attr(SubSelectableHideOutlineAttribute) !== "true");
        this.updateOutlinesFromSubSelectionElements(elementsToShow);
        if (!suppressRender) {
            this.renderOutlines();
        }
    }

    public refreshOutlines(): void {
        this.updateOutlinesFromSubSelections(this.subSelections, true /*clearExistingOutlines*/, false /*suppressRender*/);
    }

    private setOutline(outline: SubSelectionRegionOutlineFragment, visibility: SubSelectionOutlineVisibility): void {
        const helperOutline: HelperSubSelectionRegionOutline = {
            ...outline,
            visibility,
            id: outline.id as SubSelectionRegionOutlineId,
        };
        this.subSelectionRegionOutlines[outline.id] = helperOutline;
    }

    public static setDataForElement(el: HTMLElement | SVGElement, data: SubSelectionElementData): void {
        el.setAttribute(SubSelectionData, JSON.stringify(data));
    }

    public static getDataForElement(el: HTMLElement | SVGElement): SubSelectionElementData {
        return el.hasAttribute(SubSelectionData) ? JSON.parse(el.getAttribute(SubSelectionData)!) : null;
    }

    public hideAllOutlines(suppressRender: boolean = false): void {
        const allOutlines = this.subSelectionRegionOutlines;
        const updatedOutlines: HelperSubSelectionRegionOutline[] = [];
        for (const outlineId in allOutlines) {
            const outline = allOutlines[outlineId];

            updatedOutlines.push({
                ...outline,
                visibility: SubSelectionOutlineVisibility.None,
            });
        }

        this.updateRegionOutlines(updatedOutlines, suppressRender);
    }

    public getRegionOutline(id: SubSelectionRegionOutlineId): HelperSubSelectionRegionOutline | undefined {
        const outlines = this.getRegionOutlines([id]);
        return outlines[0];
    }

    public getRegionOutlines(ids: SubSelectionRegionOutlineId[]): (HelperSubSelectionRegionOutline | undefined)[] {
        return ids.map(id => this.subSelectionRegionOutlines[id]);
    }

    public getAllOutlines(): Record<string, HelperSubSelectionRegionOutline> {
        return { ...this.subSelectionRegionOutlines };
    }

    private renderOutlines() {
        const regionOutlines: HelperSubSelectionRegionOutline[] = [];
        if (this.subSelectionRegionOutlines) {
            for (const key in this.subSelectionRegionOutlines) {
                regionOutlines.push(this.subSelectionRegionOutlines[key]);
            }
        }
        this.subSelectionService.updateRegionOutlines(regionOutlines);
    }

    private getElementRegionOutlineId(selection: Selection<HTMLElement, unknown, any, unknown>): SubSelectionRegionOutlineId {
        let outlineId = selection.attr(SubSelectableObjectNameAttribute);

        let key = "";
        if (this.selectionIdCallback) {
            const selectionId: ISelectionId = this.selectionIdCallback(selection.node() as HTMLElement);
            if (selectionId?.getSelector()) {
                key = selectionId.getKey();
            }
        }

        if (key !== "") {
            return `${outlineId}___${key}` as SubSelectionRegionOutlineId;
        }
        return outlineId as SubSelectionRegionOutlineId;
    }

    public getAllSubSelectables(filterType?: SubSelectionStylesType): CustomVisualSubSelection[] | undefined {
        const subSelectables = this.getSubSelectableElements();
        const uniquenessCallback = ((a: HTMLElement, b: HTMLElement): boolean =>
            a.getAttribute(SubSelectableObjectNameAttribute) === b.getAttribute(SubSelectableObjectNameAttribute)
            && (
                !this.selectionIdCallback
                || equalsSelectionId(this.selectionIdCallback(a), this.selectionIdCallback(b))
            )
        );
        const subSelectableElements = getUniques(subSelectables, uniquenessCallback);
        let filteredTypeSubSelectableElements = subSelectableElements;
        if (filterType) {
            filteredTypeSubSelectableElements = subSelectableElements.filter((subSelectableElement: HTMLElement) => {
                const type = this.getSubSelectionTypeFromElement(subSelectableElement);
                return Number(type) === filterType;
            });
        }

        const selectionOrigins: IPoint[] = filteredTypeSubSelectableElements.map(element => {
            const boundingBox = element.getBoundingClientRect();
            return {
                x: boundingBox.x + boundingBox.width / 2,
                y: boundingBox.y + boundingBox.height / 2,
            };
        });

        const compareByY = (index1: number, index2: number) => {
            return selectionOrigins[index1].y - selectionOrigins[index2].y
        }
        const YorderedIndices: number[] = Array.from(Array(selectionOrigins.length).keys()).sort(compareByY);
        // Take all of the subselectableElements and then create visual subselections and convert into visualSubSelection[]
        const visualSubSelections: CustomVisualSubSelection[] = YorderedIndices.map(index => (
            this.createSubSelectionFromElement(
                filteredTypeSubSelectableElements[index],
                false /*showUI*/,
                undefined /*event*/,
                selectionOrigins[index],
            )
        ));
        return visualSubSelections;
    }

    public createVisualSubSelectionForSingleObject(createVisualSubSelectionArgs: CreateVisualSubSelectionFromObjectArgs): CustomVisualSubSelection {
        const { objectName, subSelectionType, displayName, showUI, selectionId, selectionOrigin, focusOrder, metadata } = createVisualSubSelectionArgs;
        const useOfssetInSelection = selectionOrigin && subSelectionType in [SubSelectionStylesType.Text, SubSelectionStylesType.NumericText];
        const origin = useOfssetInSelection ? { ...selectionOrigin, offset: { x: 0, y: (selectionOrigin?.y) * -1 } } : selectionOrigin;
        const visualSubSelection: CustomVisualSubSelection = {
            customVisualObjects: [{ objectName, selectionId: selectionId ?? undefined }],
            showUI,
            displayName,
            subSelectionType,
            selectionOrigin: origin,
            ...metadata ? { metadata } : {},
            ...focusOrder ? { focusOrder } : {},

        };
        return visualSubSelection;
    }

    private createSubSelectionFromElement(element: HTMLElement, showUI: boolean, event: PointerEvent | undefined, prevSelectionOrigin: IPoint,): CustomVisualSubSelection {
        // Need to get display name from jqdata, get selector from datum
        const objectName = element.getAttribute(SubSelectableObjectNameAttribute);

        let selectionId: powerbi.visuals.ISelectionId;
        if (this.selectionIdCallback) {
            selectionId = this.selectionIdCallback(element);
        }

        const subSelectionType = this.getSubSelectionTypeFromElement(element);
        const displayName = this.getDisplayNameFromElement(element);
        let selectionOrigin: IPoint = prevSelectionOrigin;
        if (event) {
            selectionOrigin = {
                x: event.clientX,
                y: event.clientY,
            };
        }

        const visualSubSelection = this.createVisualSubSelectionForSingleObject({
            objectName,
            subSelectionType,
            displayName,
            showUI,
            selectionId,
            selectionOrigin,
        });
        return visualSubSelection;
    }

    private getDisplayNameFromElement(element: HTMLElement): string {
        return element.getAttribute(SubSelectableDisplayNameAttribute) ?? '';
    }

    private getSubSelectionTypeFromElement(element: HTMLElement): SubSelectionStylesType | undefined {
        const type = element.getAttribute(SubSelectableTypeAttribute);
        if (!type) {
            return undefined;
        }
        return Number(type);
    }

    private getSubSelectableElements(): HTMLElement[] {
        const hostElement = this.hostElement;
        return this.host
            .selectAll<HTMLElement, unknown>(HtmlSubSelectableSelector)
            .filter(function () {
                const element = this;
                return helperOwnsElement(hostElement, element);
            }).nodes();
    }
}
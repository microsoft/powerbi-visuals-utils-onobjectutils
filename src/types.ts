import SubSelectionOutlineVisibility = powerbi.visuals.SubSelectionOutlineVisibility;
import SubSelectionRegionOutlineId = powerbi.visuals.SubSelectionRegionOutlineId;
import SubSelectionRegionOutline = powerbi.visuals.SubSelectionRegionOutline;
import CustomVisualSubSelection = powerbi.visuals.CustomVisualSubSelection;
import SubSelectionOutlineRestrictionOptions = powerbi.visuals.SubSelectionOutlineRestrictionOptions;
import ISelectionId = powerbi.visuals.ISelectionId;
import SubSelectionRegionOutlineFragment = powerbi.visuals.SubSelectionRegionOutlineFragment;

export interface SubSelectionElementData {
    outlineRestrictionOptions?: SubSelectionOutlineRestrictionOptions;
}

export interface HtmlSubSelectionSource {
    subSelectionElement: HTMLElement;
    visualSubSelection: CustomVisualSubSelection;
}

export interface HtmlSubselectionHelperArgs {
    /** Element which contains the items that can be sub-selected */
    hostElement: HTMLElement;
    subSelectionService: powerbi.extensibility.IVisualSubSelectionService;
    selectionIdCallback?: ((e: HTMLElement) => ISelectionId);
    customOutlineCallback?: ((subSelection: CustomVisualSubSelection) => SubSelectionRegionOutlineFragment[]);
    customElementCallback?: ((subSelection: CustomVisualSubSelection) => HTMLElement[]);
    subSelectionMetadataCallback?: ((subSelectionElement: HTMLElement) => unknown);
}

export interface CreateVisualSubSelectionFromObjectArgs {
    objectName: string;
    subSelectionType: powerbi.visuals.SubSelectionStylesType;
    displayName: string;
    showUI: boolean;
    selectionId?: ISelectionId;
    selectionOrigin?: powerbi.extensibility.IPoint;
    focusOrder?: number;
    metadata?: unknown;
}

/**
 * Helper tool that makes it easier for visuals to manage sub-selection.
 * Provides methods for easily making elements sub-selectable, abstracting away most of the complexity around it
 * For a full implementation, see the HTMLSubSelectionHelper
 */
export interface ISubSelectionHelper<TElement, TIdentifier = unknown> {
    destroy(): void;
    setFormatMode(isFormatMode: boolean): void;

    /**
     * Updates the outline for the given element. Will use a region with Rectangle outlines, creating it if needed
     * @returns the id of the outline
     */
    updateElementOutline(
        element: TElement,
        visibility: SubSelectionOutlineVisibility,
    ): SubSelectionRegionOutlineId;

    /**
     * Updates the outlines for the given elements.  Will use a region with Rectangle outlines, creating if needed
     * @returns the ids of the created outlines
     */
    updateElementOutlines(
        elements: TElement[],
        visibility: SubSelectionOutlineVisibility,
    ): SubSelectionRegionOutlineId[];

    /**
     * Updates the given outline
     */
    updateRegionOutline(
        outline: SubSelectionRegionOutline,
    ): void;

    /**
     * Updates the given outlines
     */
    updateRegionOutlines(
        outlines: SubSelectionRegionOutline[],
    ): void;

    /**
     * Gets the outline for the id
     */
    getRegionOutline(id: SubSelectionRegionOutlineId): SubSelectionRegionOutline | undefined;

    /**
     * Gets the outlines for the ids
     */
    getRegionOutlines(id: SubSelectionRegionOutlineId[]): (SubSelectionRegionOutline | undefined)[];

    /**
     * Gets all currently available subselectable elements
     */
    getAllSubSelectables(): TIdentifier[];

    /**
     * Allows creation of custom visual subselections that don't have a DOMElement
     */
    createVisualSubSelectionForSingleObject(createVisualSubSelectionArgs: CreateVisualSubSelectionFromObjectArgs): TIdentifier;
}
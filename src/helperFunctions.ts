type Comparator<T> = (a: T, b: T) => boolean;

export function getObjectValues<T>(obj: Record<string, T>): T[] {
    if (!obj) {
        return [];
    }

    return Object.keys(obj).map(key => obj[key])
}

export function groupArrayElements(array: HTMLElement[], func: (element: HTMLElement) => powerbi.visuals.SubSelectionRegionOutlineId): Record<string, HTMLElement[]> {
    if (!array) {
        return {};
    }
    // Use reduce to iterate over the array and accumulate an object
    return array.reduce((acc, cur) => {
        // Get the key by applying the function to the current element
        const key = func(cur);
        // If the key already exists in the accumulator, push the current element to the array
        if (acc[key]) {
            acc[key].push(cur);
        } else {
            // Otherwise, create a new array with the current element
            acc[key] = [cur];
        }
        // Return the accumulator
        return acc;
    }, {});
}

export function isEqual(value: any, other: any): boolean {
    // Check if the values are strictly equal
    if (value === other) {
        return true;
    }

    // Check if both values are objects
    if (value && other && typeof value === 'object' && typeof other === 'object') {
        const keysA = Object.keys(value);
        const keysB = Object.keys(other);

        // Check if the objects have the same number of properties
        if (keysA.length !== keysB.length) {
            return false;
        }

        // Check if all properties are equal
        for (const key of keysA) {
            if (!isEqual(value[key], other[key])) {
                return false;
            }
        }
        return true;
    }
    return false;
}

export function isArrayEmpty(array: any[]): boolean {
    if (!array || array.length === 0) {
        return true;
    }
    return false;
}

export function getUniques<T>(array: T[], comparator: Comparator<T>): T[] {
    if (!array) {
        return []
    }
    return array.reduce((result, current) => {
        if (!result.some(item => comparator(item, current))) {
            result.push(current);
        }
        return result;
    }, [] as T[]);
}

export function debounce(func: any, delay: any) {
    let timeoutId;

    return function (...args) {
        const context = this;

        clearTimeout(timeoutId);

        timeoutId = setTimeout(() => {
            func.apply(context, args);
        }, delay);
    };
}

export function equalsSelectionId(x: powerbi.visuals.ISelectionId, y: powerbi.visuals.ISelectionId): boolean {
    // Normalize falsy to null
    x = x || null;
    y = y || null;

    if (x === y)
        return true;

    if (!x !== !y)
        return false;

    return x.equals(y) && y.equals(x);
}
declare module 'pptx-preview' {
    export interface PptxPreviewOptions {
        width: number;
        height?: number;
        mode?: 'slide' | 'scroll';
    }

    export interface PptxPreviewInstance {
        preview(data: ArrayBuffer): Promise<unknown>;
        load?(data: ArrayBuffer): Promise<unknown>;
        destroy?(): void;
    }

    export function init(container: HTMLElement, options: PptxPreviewOptions): PptxPreviewInstance;
}


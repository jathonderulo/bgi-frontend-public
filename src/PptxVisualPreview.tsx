import {useEffect, useRef, useState} from 'react';
import {init} from 'pptx-preview';

type Props = {
    arrayBuffer: ArrayBuffer;
};

type PptxPreviewInstance = {
    preview: (data: ArrayBuffer) => Promise<unknown>;
    destroy?: () => void;
};

export function PptxVisualPreview({arrayBuffer}: Props) {
    const hostRef = useRef<HTMLDivElement | null>(null);
    const [error, setError] = useState<string | null>(null);

    useEffect(() => {
        let disposed = false;
        let previewInstance: PptxPreviewInstance | null = null;

        const run = async () => {
            if (!hostRef.current) return;
            hostRef.current.innerHTML = '';
            setError(null);

            try {
                const width = Math.max(720, hostRef.current.clientWidth || 720);
                const height = Math.max(520, hostRef.current.clientHeight || 520);
                previewInstance = init(hostRef.current, { width, height, mode: 'scroll' }) as PptxPreviewInstance;
                await previewInstance.preview(arrayBuffer);
            } catch (err) {
                console.error('PPTX visual preview failed:', err);
                if (!disposed) {
                    setError('Could not render this PPTX visually. Please use Download.');
                }
            }
        };

        run();

        return () => {
            disposed = true;
            if (previewInstance?.destroy) {
                previewInstance.destroy();
            }
            if (hostRef.current) {
                hostRef.current.innerHTML = '';
            }
        };
    }, [arrayBuffer]);

    return (
        <div style={{height: '100%', background: '#eef1f4'}}>
            {error && (
                <div style={{padding: 12, color: '#a33', fontSize: 13}}>
                    {error}
                </div>
            )}
            <div ref={hostRef} style={{height: '100%', overflow: 'auto'}} />
        </div>
    );
}


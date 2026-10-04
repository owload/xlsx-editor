/**
 * The few drawing calls the preview needs, so the layout can be tested without a canvas and the
 * drawing code does not depend on where the canvas comes from.
 */
export interface DrawContext {
  fillStyle: string | CanvasGradient | CanvasPattern;
  font: string;
  textBaseline: string;
  textAlign: string;
  fillRect(x: number, y: number, width: number, height: number): void;
  fillText(text: string, x: number, y: number): void;
  scale(x: number, y: number): void;
}

export interface PngCanvas {
  context: DrawContext;
  /** The bytes of the picture drawn so far, as a PNG. */
  toPng(): Promise<Uint8Array>;
}

export type CanvasFactory = (width: number, height: number) => PngCanvas | null;

/** An OffscreenCanvas, or null where the environment has none (then there is simply no preview). */
export const browserCanvas: CanvasFactory = (width, height) => {
  if (typeof OffscreenCanvas === 'undefined') return null;
  const canvas = new OffscreenCanvas(width, height);
  const context = canvas.getContext('2d');
  if (!context) return null;
  return {
    context,
    toPng: async () => new Uint8Array(await (await canvas.convertToBlob({ type: 'image/png' })).arrayBuffer()),
  };
};

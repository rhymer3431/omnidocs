declare module 'mammoth/mammoth.browser' {
  interface MammothMessage {
    type: string;
    message: string;
  }

  interface MammothResult {
    value: string;
    messages: MammothMessage[];
  }

  interface MammothParagraph {
    type: 'paragraph';
    styleId: string | null;
    styleName: string | null;
    alignment: string | null;
    numbering: unknown;
    indent: { start?: string | null; end?: string | null; firstLine?: string | null; hanging?: string | null };
    [key: string]: unknown;
  }

  interface MammothRun {
    type: 'run';
    styleId: string | null;
    styleName: string | null;
    font: string | null;
    fontSize: number | null;
    highlight: string | null;
    [key: string]: unknown;
  }

  interface MammothOptions {
    styleMap?: string[];
    transformDocument?: (document: unknown) => unknown;
  }

  const mammoth: {
    convertToHtml(input: { arrayBuffer: ArrayBuffer }, options?: MammothOptions): Promise<MammothResult>;
    transforms: {
      paragraph(transform: (paragraph: MammothParagraph) => MammothParagraph): (document: unknown) => unknown;
      run(transform: (run: MammothRun) => MammothRun): (document: unknown) => unknown;
    };
  };

  export default mammoth;
}

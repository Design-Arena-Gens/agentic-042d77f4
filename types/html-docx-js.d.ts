declare module "html-docx-js" {
  type HtmlDocx = {
    asBlob: (html: string, options?: Record<string, unknown>) => Blob;
  };

  const htmlDocx: HtmlDocx;
  export default htmlDocx;
}

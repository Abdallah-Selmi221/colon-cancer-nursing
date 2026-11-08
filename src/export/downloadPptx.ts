/**
 * Triggers the download of the validated standalone PowerPoint (.pptx) presentation.
 */
export function downloadPptxPresentation(): void {
  const base = import.meta.env.BASE_URL || '/';
  const fileUrl = `${base}colon-cancer-presentation.pptx`;
  
  const link = document.createElement('a');
  link.href = fileUrl;
  link.download = 'colon-cancer-presentation.pptx';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

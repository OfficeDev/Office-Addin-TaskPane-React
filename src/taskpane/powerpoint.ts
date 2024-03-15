/* global PowerPoint console */

export async function insertText(text: string) {
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const textBox = slide.shapes.addTextBox(text);
      textBox.fill.setSolidColor("white");
      textBox.lineFormat.color = "black";
      textBox.lineFormat.weight = 1;
      textBox.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.solid;
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

<template>
  <span>The following example demonstrates how to add shape group in a Word document</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from "vue";

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);

        // Create a new document object  
        let doc = wasmModule.Document.Create();
        let sec = doc.AddSection();

        // Add a new paragraph
        let para = sec.AddParagraph();
        // Add a shape group with the height and width
        let shapegroup = para.AppendShapeGroup(375, 462);
        shapegroup.HorizontalPosition = 180;

        //calcuate the scale ratio
        let X = (shapegroup.Width / 1000.0);
        let Y = (shapegroup.Height / 1000.0);

        // Create a textbox
        let txtBox = wasmModule.TextBox.Create(doc);
        txtBox.SetShapeType(wasmModule.ShapeType.RoundRectangle);
        txtBox.Width = 125 / X;
        txtBox.Height = 54 / Y;
        let paragraph = txtBox.Body.AddParagraph();
        paragraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
        paragraph.AppendText("Start");
        txtBox.HorizontalPosition = 19 / X;
        txtBox.VerticalPosition = 27 / Y;
        txtBox.Format.LineColor = wasmModule.Color.get_Green();
        shapegroup.ChildObjects.Add(txtBox);

        // Create an arrow line shape
        let arrowLineShape = wasmModule.ShapeObject.Create(doc, wasmModule.ShapeType.DownArrow);
        arrowLineShape.Width = 16 / X;
        arrowLineShape.Height = 40 / Y;
        arrowLineShape.HorizontalPosition = 69 / X;
        arrowLineShape.VerticalPosition = 87 / Y;
        arrowLineShape.StrokeColor = wasmModule.Color.get_Purple();
        shapegroup.ChildObjects.Add(arrowLineShape);

        // Create a textbox
        txtBox = wasmModule.TextBox.Create(doc);
        txtBox.SetShapeType(wasmModule.ShapeType.Rectangle);
        txtBox.Width = 125 / X;
        txtBox.Height = 54 / Y;
        paragraph = txtBox.Body.AddParagraph();
        paragraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
        paragraph.AppendText("Step 1");
        txtBox.HorizontalPosition = 19 / X;
        txtBox.VerticalPosition = 131 / Y;
        txtBox.Format.LineColor = wasmModule.Color.get_Blue();
        shapegroup.ChildObjects.Add(txtBox);

        // Create an arrow line shape
        arrowLineShape = wasmModule.ShapeObject.Create(doc, wasmModule.ShapeType.DownArrow);
        arrowLineShape.Width = 16 / X;
        arrowLineShape.Height = 40 / Y;
        arrowLineShape.HorizontalPosition = 69 / X;
        arrowLineShape.VerticalPosition = 192 / Y;
        arrowLineShape.StrokeColor = wasmModule.Color.get_Purple();
        shapegroup.ChildObjects.Add(arrowLineShape);

        // Create an arrow line shape
        txtBox = wasmModule.TextBox.Create(doc);
        txtBox.SetShapeType(wasmModule.ShapeType.Parallelogram);
        txtBox.Width = 149 / X;
        txtBox.Height = 59 / Y;
        paragraph = txtBox.Body.AddParagraph();
        paragraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
        paragraph.AppendText("Step 2");
        txtBox.HorizontalPosition = 7 / X;
        txtBox.VerticalPosition = 236 / Y;
        txtBox.Format.LineColor = wasmModule.Color.get_BlueViolet();
        shapegroup.ChildObjects.Add(txtBox);

        // Create an arrow line shape
        arrowLineShape = wasmModule.ShapeObject.Create(doc, wasmModule.ShapeType.DownArrow);
        arrowLineShape.Width = 16 / X;
        arrowLineShape.Height = 40 / Y;
        arrowLineShape.HorizontalPosition = 66 / X;
        arrowLineShape.VerticalPosition = 300 / Y;
        arrowLineShape.StrokeColor = wasmModule.Color.get_Purple();
        shapegroup.ChildObjects.Add(arrowLineShape);

        // Create an arrow line shape
        txtBox = wasmModule.TextBox.Create(doc);
        txtBox.SetShapeType(wasmModule.ShapeType.Rectangle);
        txtBox.Width = 125 / X;
        txtBox.Height = 54 / Y;
        paragraph = txtBox.Body.AddParagraph();
        paragraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
        paragraph.AppendText("Step 3");
        txtBox.HorizontalPosition = 19 / X;
        txtBox.VerticalPosition = 345 / Y;
        txtBox.Format.LineColor = wasmModule.Color.get_Blue();
        shapegroup.ChildObjects.Add(txtBox);

        // Save the document
        const outputFileName = "AddShapeGroup.docx";
        doc.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2010 });

        // Read the saved document from the virtual file system and convert it to a byte array
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);

        // Convert the byte array into a Blob object
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

        // Dispose of the document object to free resources
        doc.Dispose();

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
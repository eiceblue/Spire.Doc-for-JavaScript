<template>
  <span>Click the following button to create a nested table in a Word document</span>
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
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF","/Library/Fonts/",`${import.meta.env.BASE_URL}static/font/`);

        //Create a new document
        let doc = wasmModule.Document.Create();
        let section = doc.AddSection();

        //Add a table
        let table = section.AddTable({showBorder: true});
        table.ResetCells(2, 2);

        //Set column width
        table.Rows.get(0).Cells.get(0).SetCellWidth(70, wasmModule.CellWidthType.Point);
        table.Rows.get(1).Cells.get(0).SetCellWidth(70, wasmModule.CellWidthType.Point);
        table.AutoFit(wasmModule.AutoFitBehaviorType.AutoFitToWindow);

        //Insert content to cells
        table.Rows.get(0).Cells.get(0).AddParagraph().AppendText("Spire.Doc for .NET");
        let text = "Spire.Doc for .NET is a professional Word" +
            ".NET library specifically designed for developers to create," +
            "read, write, convert and print Word document files from any .NET" +
            "platform with fast and high quality performance.";
        table.Rows.get(0).Cells.get(1).AddParagraph().AppendText(text);

        //Add a nested table to cell(first row, second column)
        let nestedTable = table.Rows.get(0).Cells.get(1).AddTable({showBorder: true});
        nestedTable.ResetCells(4, 3);
        nestedTable.AutoFit(wasmModule.AutoFitBehaviorType.AutoFitToContents);

        //Add content to nested cells
        nestedTable.Rows.get(0).Cells.get(0).AddParagraph().AppendText("NO.");
        nestedTable.Rows.get(0).Cells.get(1).AddParagraph().AppendText("Item");
        nestedTable.Rows.get(0).Cells.get(2).AddParagraph().AppendText("Price");

        nestedTable.Rows.get(1).Cells.get(0).AddParagraph().AppendText("1");
        nestedTable.Rows.get(1).Cells.get(1).AddParagraph().AppendText("Pro Edition");
        nestedTable.Rows.get(1).Cells.get(2).AddParagraph().AppendText("$799");

        nestedTable.Rows.get(2).Cells.get(0).AddParagraph().AppendText("2");
        nestedTable.Rows.get(2).Cells.get(1).AddParagraph().AppendText("Standard Edition");
        nestedTable.Rows.get(2).Cells.get(2).AddParagraph().AppendText("$599");

        nestedTable.Rows.get(3).Cells.get(0).AddParagraph().AppendText("3");
        nestedTable.Rows.get(3).Cells.get(1).AddParagraph().AppendText("Free Edition");
        nestedTable.Rows.get(3).Cells.get(2).AddParagraph().AppendText("$0");


        // Define the output file name
        const outputFileName = "CreateNestedTable_output.docx";

        // Save the document to the specified path
        doc.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Clean up resources
        doc.Close();
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

<template>
  <span>The following example shows how to set text direction in a Word document. </span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName"> Click here to download the generated file </a>
</template>

<script>
import { ref } from 'vue';

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref('');

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

        //Create a new document
        let doc = wasmModule.Document.Create();

        //Add the first section
        let section1 = doc.AddSection();
        //Set text direction for all text in a section
        section1.TextDirection = wasmModule.TextDirection.RightToLeft;

        //Set Font Style and Size
        let style = wasmModule.ParagraphStyle.Create(doc);
        style.Name = 'FontStyle';
        style.CharacterFormat.FontName = 'Arial';
        style.CharacterFormat.FontSize = 15;

        doc.Styles.Add(style);

        //Add two paragraphs and apply the font style
        let p = section1.AddParagraph();
        p.AppendText('Only Spire.Doc, no Microsoft Office automation');
        p.ApplyStyle({ styleName: style.Name });
        p = section1.AddParagraph();
        p.AppendText('Convert file documents with high quality');
        p.ApplyStyle({ styleName: style.Name });

        //Set text direction for a part of text
        //Add the second section
        let section2 = doc.AddSection();
        //Add a table
        let table = section2.AddTable();
        table.ResetCells(1, 1);
        let cell = table.Rows.get(0).Cells.get(0);
        table.Rows.get(0).Height = 150;
        table.Rows.get(0).Cells.get(0).SetCellWidth(10, wasmModule.CellWidthType.Point);
        //Set vertical text direction of table
        cell.CellFormat.TextDirection = wasmModule.TextDirection.RightToLeftRotated;
        cell.AddParagraph().AppendText('This is vertical style');
        //Add a paragraph and set horizontal text direction
        p = section2.AddParagraph();
        p.AppendText('This is horizontal style');
        p.ApplyStyle(style.Name);
        // Define the output file name
        const outputFileName = 'SetTextDirection.docx';

        // Save the document to the specified path
        doc.SaveToFile({
          fileName: outputFileName,
          fileFormat: wasmModule.FileFormat.Docx2013,
        });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });

        // Clean up resources
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

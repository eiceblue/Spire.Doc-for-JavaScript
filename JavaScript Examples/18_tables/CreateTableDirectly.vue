<template>
  <span>Click the following button to directly creat a table in a Word document</span>
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

        //Create a Word document
        let doc = wasmModule.Document.Create();

        //Add a section
        let section = doc.AddSection();

        //Create a table
        let table = wasmModule.Table.Create(doc, false);
        //Set the width of table
        table.PreferredWidth = wasmModule.PreferredWidth.Create(wasmModule.WidthType.Percentage, 100);
        //Set the border of table
        table.TableFormat.Borders.BorderType = wasmModule.BorderStyle.Single;

        //Create a table row
        let row = wasmModule.TableRow.Create(doc, false);
        row.Height = 50.0;
        table.Rows.Add(row);

        //Create a table cell
        let cell1 = wasmModule.TableCell.Create(doc);
        //Add a paragraph
        let para1 = cell1.AddParagraph();
        //Append text in the paragraph
        para1.AppendText("Row 1, Cell 1");
        //Set the horizontal alignment of paragrah
        para1.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
        //Set the background color of cell
        cell1.CellFormat.BackColor = wasmModule.Color.get_CadetBlue();
        //Set the vertical alignment of paragraph
        cell1.CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
        row.Cells.Add(cell1);

        //Create a table cell
        let cell2 = wasmModule.TableCell.Create(doc);
        let para2 = cell2.AddParagraph();
        para2.AppendText("Row 1, Cell 2");
        para2.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
        cell2.CellFormat.BackColor = wasmModule.Color.get_CadetBlue();
        cell2.CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
        row.Cells.Add(cell2);

        //Add the table in the section
        section.Tables.Add(table);

        // Define the output file name
        const outputFileName = "CreateTableDirectly_output.docx";

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

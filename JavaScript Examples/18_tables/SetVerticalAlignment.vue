<template>
  <span>Click the following button to set vertical alignment for table in a Word document</span>
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

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "E-iceblue.png";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`); 

        //Create a new Word document and add a new section
        let doc = wasmModule.Document.Create();
        let section = doc.AddSection();

        //Add a table with 3 columns and 3 rows
        let table = section.AddTable({showBorder: true});
        table.ResetCells(3, 3);

        //Merge rows
        table.ApplyVerticalMerge(0, 0, 2);

        //Set the vertical alignment for each cell, default is top
        table.Rows.get(0).Cells.get(0).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
        table.Rows.get(0).Cells.get(1).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Top;
        table.Rows.get(0).Cells.get(2).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Top;
        table.Rows.get(1).Cells.get(1).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
        table.Rows.get(1).Cells.get(2).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle
        table.Rows.get(2).Cells.get(1).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Bottom;
        table.Rows.get(2).Cells.get(2).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Bottom;

        //Inset a picture into the table cell
        let paraPic = table.Rows.get(0).Cells.get(0).AddParagraph();

        let pic = paraPic.AppendPicture({imgFile: inputFileName});

        //Create data
        let data = [
            ["", "Spire.Office", "Spire.DataExport"],
            ["", "Spire.Doc", "Spire.DocViewer"],
            ["", "Spire.XLS", "Spire.PDF"]
        ];

        //Append data to table
        for (let r = 0; r < 3; r++) {
            let dataRow = table.Rows.get(r);
            dataRow.Height = 50;
            for (let c = 0; c < 3; c++) {
                if (c == 1) {
                    let par = dataRow.Cells.get(c).AddParagraph();
                    par.AppendText(data[r][c]);
                    dataRow.Cells.get(c).Width = (section.PageSetup.ClientWidth) / 2;
                }
                if (c == 2) {
                    let par = dataRow.Cells.get(c).AddParagraph();
                    par.AppendText(data[r][c]);
                    dataRow.Cells.get(c).Width = (section.PageSetup.ClientWidth) / 2;
                }
            }
        }

        // Define the output file name
        const outputFileName = "SetVerticalAlignment_output.docx";

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

<template>
  <span>Click the following button to repeat row on each page</span>
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

        //Create word document
        let document = wasmModule.Document.Create();

        //Create a new section
        let section = document.AddSection();

        //Create a table width default borders
        let table = section.AddTable({showBorder: true});
        //Set table with to 100%
        let width = wasmModule.PreferredWidth.Create(wasmModule.WidthType.Percentage, 100);
        table.PreferredWidth = width;

        //Add a new row
        let row = table.AddRow();
        //Set the row as a table header
        row.IsHeader = true;
        //Set the backcolor of row
        row.RowFormat.BackColor = wasmModule.Color.get_LightGray();
        //Add a new cell for row
        let cell = row.AddCell();
        cell.SetCellWidth(100, wasmModule.CellWidthType.Percentage);
        //Add a paragraph for cell to put some data
        let parapraph = cell.AddParagraph();
        //Add text
        parapraph.AppendText("Row Header 1");
        //Set paragraph horizontal center alignment
        parapraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;

        row = table.AddRow({isCopyFormat: false, columnsNum: 1});
        row.IsHeader = true;
        row.RowFormat.BackColor = wasmModule.Color.get_Ivory();
        //Set row height
        row.Height = 30;
        cell = row.Cells.get(0);
        cell.SetCellWidth(100, wasmModule.CellWidthType.Percentage);
        //Set cell vertical middle alignment
        cell.CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
        //Add a paragraph for cell to put some data
        parapraph = cell.AddParagraph();
        //Add text
        parapraph.AppendText("Row Header 2");
        parapraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;

        //Add many common rows
        for (let i = 0; i < 70; i++) {
            row = table.AddRow({isCopyFormat: false, columnsNum: 2});
            cell = row.Cells.get(0);
            //Set cell width to 50% of table width
            cell.SetCellWidth(50, wasmModule.CellWidthType.Percentage);
            cell.AddParagraph().AppendText("Column 1 Text");
            cell = row.Cells.get(1);
            cell.SetCellWidth(50, wasmModule.CellWidthType.Percentage);
            cell.AddParagraph().AppendText("Column 2 Text");
        }
        //Set cell backcolor
        for (let j = 1; j < table.Rows.Count; j++) {
            if (j % 2 == 0) {
                let row2 = table.Rows.get(j);
                for (let f = 0; f < row2.Cells.Count; f++) {
                    row2.Cells.get(f).CellFormat.BackColor = wasmModule.Color.get_LightBlue();
                }
            }
        }

        // Define the output file name
        const outputFileName = "RepeatRowOnEachPage_output.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Clean up resources
        document.Close();
        document.Dispose();

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

<template>
  <span>Click the following button to get diagonal border of table cell</span>
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
        let inputFileName = "GetDiagonalBorderOfCell.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);  

        //Load Word from disk
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Get the first section
        let section = doc.Sections.get(0);

        //Get the first table in the section
        let table = section.Tables.get_Item(0);

        let stringBuilder = [];

        //Get the setting of the diagonal border of table cell
        let bs_UP = table.Rows.get(0).Cells.get(0).CellFormat.Borders.DiagonalUp.BorderType;
        stringBuilder.push("DiagonalUp border type of table cell(0,0) is " + bs_UP + "\n");
        let color_UP = table.Rows.get(0).Cells.get(0).CellFormat.Borders.DiagonalUp.Color;
        stringBuilder.push("DiagonalUp border color of table cell(0,0) is " + color_UP.ToString() + "\n");
        let width_UP = table.Rows.get(0).Cells.get(0).CellFormat.Borders.DiagonalUp.LineWidth;
        stringBuilder.push("Line width of DiagonalUp border of table cell(0,0) is " + width_UP + "\n");
        let bs_Down = table.Rows.get(0).Cells.get(0).CellFormat.Borders.DiagonalDown.BorderType;
        stringBuilder.push("DiagonalDown border type of table cell(0,0) is " + bs_Down + "\n");
        let color_Down = table.Rows.get(0).Cells.get(0).CellFormat.Borders.DiagonalDown.Color;
        stringBuilder.push("DiagonalDown border color of table cell(0,0) is " + color_Down.ToString() + "\n");
        let width_Down = table.Rows.get(0).Cells.get(0).CellFormat.Borders.DiagonalDown.LineWidth;
        stringBuilder.push("DiagonalDown border color of table cell(0,0) is " + width_Down + "\n");

        // Define the output file name
        const outputFileName = "GetDiagonalBorder_output.txt";

        // Save the string to the specified path
        FS.writeFile(outputFileName, stringBuilder.join('\n')); 

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "text/plain"});

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

<template>
  <span>Click the following button to modify table format including row format and cell format</span>
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
        let inputFileName = "ModifyTableFormat.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`); 

        //Load Word document from disk
        let document = wasmModule.Document.Create();
        document.LoadFromFile(inputFileName);

        //Get the first section
        let section = document.Sections.get(0);

        //Get tables
        let tb1 = section.Tables.get_Item(0);
        let tb2 = section.Tables.get_Item(1);
        let tb3 = section.Tables.get_Item(2);

        MoidyTableFormat(tb1);
        ModifyRowFormat(tb2);
        ModifyCellFormat(tb3);

        // Define the output file name
        const outputFileName = "ModifyTableFormat_output.docx";

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

      function MoidyTableFormat(table) {
    //Set table width
    table.PreferredWidth = wasmModule.PreferredWidth.Create(wasmModule.WidthType.Twip, 6000);

    //Apply style for table
    table.ApplyStyle(wasmModule.DefaultTableStyle.ColorfulGridAccent3);

    //Set table padding
    table.TableFormat.Paddings.All = 5;

    //Set table title and description
    table.Title = "Spire.Doc for .NET";
    table.TableDescription = "Spire.Doc for .NET is a professional Word .NET library";
}
function ModifyRowFormat(table) {
    //Set cell spacing
    table.Rows.get(0).RowFormat.CellSpacing = 2;

    //Set row height
    table.Rows.get(1).HeightType = wasmModule.TableRowHeightType.Exactly;
    table.Rows.get(1).Height = 20;

    //Set background color
    table.Rows.get(2).RowFormat.BackColor = wasmModule.Color.get_DarkSeaGreen();

}
function ModifyCellFormat(table) {
    //Set alignment
    table.Rows.get(0).Cells.get(0).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
    table.Rows.get(0).Cells.get(0).Paragraphs.get_Item(0).Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;

    //Set background color
    table.Rows.get(1).Cells.get(0).CellFormat.BackColor = wasmModule.Color.get_DarkSeaGreen();

    //Set cell border
    table.Rows.get(2).Cells.get(0).CellFormat.Borders.BorderType = wasmModule.BorderStyle.Single;
    table.Rows.get(2).Cells.get(0).CellFormat.Borders.LineWidth = 1;
    table.Rows.get(2).Cells.get(0).CellFormat.Borders.Left.Color = wasmModule.Color.get_Red();
    table.Rows.get(2).Cells.get(0).CellFormat.Borders.Right.Color = wasmModule.Color.get_Red();
    table.Rows.get(2).Cells.get(0).CellFormat.Borders.Top.Color = wasmModule.Color.get_Red();
    table.Rows.get(2).Cells.get(0).CellFormat.Borders.Bottom.Color = wasmModule.Color.get_Red();

    //Set text direction
    table.Rows.get(3).Cells.get(0).CellFormat.TextDirection = wasmModule.TextDirection.RightToLeft;
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

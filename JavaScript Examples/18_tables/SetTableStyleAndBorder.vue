<template>
  <span>Click the following button to set the style of table in a Word document</span>
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
        let inputFileName = "TableSample.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`); 

        //Create a document and load file
        let document = wasmModule.Document.Create();
        document.LoadFromFile(inputFileName);

        let section = document.Sections.get(0);

        //Get the first table
        let table = section.Tables.get_Item(0);

        //Apply the table style
        table.ApplyStyle(wasmModule.DefaultTableStyle.ColorfulList);

        //Set right border of table
        table.TableFormat.Borders.Right.BorderType = wasmModule.BorderStyle.Hairline;
        table.TableFormat.Borders.Right.LineWidth = 1.0;
        table.TableFormat.Borders.Right.Color = wasmModule.Color.get_Red();

        //Set top border of table
        table.TableFormat.Borders.Top.BorderType = wasmModule.BorderStyle.Hairline;
        table.TableFormat.Borders.Top.LineWidth = 1.0;
        table.TableFormat.Borders.Top.Color = wasmModule.Color.get_Green();

        //Set left border of table
        table.TableFormat.Borders.Left.BorderType = wasmModule.BorderStyle.Hairline;
        table.TableFormat.Borders.Left.LineWidth = 1.0;
        table.TableFormat.Borders.Left.Color = wasmModule.Color.get_Yellow();

        //Set bottom border is none
        table.TableFormat.Borders.Bottom.BorderType = wasmModule.BorderStyle.DotDash;

        //Set vertical and horizontal border
        table.TableFormat.Borders.Vertical.BorderType = wasmModule.BorderStyle.Dot;
        table.TableFormat.Borders.Horizontal.BorderType = wasmModule.BorderStyle.None;
        table.TableFormat.Borders.Vertical.Color = wasmModule.Color.get_Orange();


        // Define the output file name
        const outputFileName = "SetTableStyleAndBorder_output.docx";

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

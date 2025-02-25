<template>
  <span>Click the following button to get table position</span>
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
        let inputFileName = "TableSample-Az.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`); 

        //Create a document
        let document = wasmModule.Document.Create();
        //Load file
        document.LoadFromFile(inputFileName);
        //Get the first section
        let section = document.Sections.get_Item(0);
        //Get the first table
        let table = section.Tables.get_Item(0);

        let stringBuidler = [];

        //Verify whether the table uses "Around" text wrapping or not.
        if (table.TableFormat.WrapTextAround) {

            let positon = table.TableFormat.Positioning;
            stringBuidler.push("Horizontal:\n");
            stringBuidler.push("Position:" + positon.HorizPosition + " pt\n");
            stringBuidler.push("Absolute Position:" + positon.HorizPositionAbs + ", Relative to:" + positon.HorizRelationTo + "\n");
            stringBuidler.push("\n");
            stringBuidler.push("Vertical:\n");
            stringBuidler.push("Position:" + positon.VertPosition + " pt\n");
            stringBuidler.push("Absolute Position:" + positon.VertPositionAbs + ", Relative to:" + positon.VertRelationTo + "\n");
            stringBuidler.push("\n");
            stringBuidler.push("Distance from surrounding text:\n");
            stringBuidler.push("Top:" + positon.DistanceFromTop + " pt, Left:" + positon.DistanceFromLeft + " pt\n");
            stringBuidler.push("Bottom:" + positon.DistanceFromBottom + "pt, Right:" + positon.DistanceFromRight + " pt\n");
        }

        // Define the output file name
        const outputFileName = "GetTablePosition_output.txt";

        // Save the document to the specified path
        FS.writeFile(outputFileName, stringBuidler.join('\n'));

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "text/plain"});

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

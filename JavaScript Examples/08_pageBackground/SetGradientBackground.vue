<template>
  <span>Click the following button to set gradient background for Word document.</span>
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

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "Template_Docx_2.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        //Create Word document.
        let document = wasmModule.Document.Create();

        //Load the file from disk.
        document.LoadFromFile(inputFileName);

        //Set the background type as Gradient.
        document.Background.Type = wasmModule.BackgroundType.Gradient;
        let Test = document.Background.Gradient;

        //Set the first color and second color for Gradient.
        Test.Color1 = wasmModule.Color.get_White();
        Test.Color2 = wasmModule.Color.get_LightBlue();

        //Set the Shading style and Variant for the gradient.
        Test.ShadingVariant = wasmModule.GradientShadingVariant.ShadingDown;
        Test.ShadingStyle = wasmModule.GradientShadingStyle.Horizontal;

        // Define the output file name
        const outputFileName = "SetGradientBackground_out.docx";

        // Save the document to the specified path
        document.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

        // Clean up resources
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

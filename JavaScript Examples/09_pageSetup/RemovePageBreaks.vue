<template>
  <span>Click the following button to remove page breaks in Word document.</span>
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
        let inputFileName = "Template_Docx_4.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        //Create Word document.
        let document = wasmModule.Document.Create();

        //Load the file from disk.
        document.LoadFromFile(inputFileName);

        //Traverse every paragraph of the first section of the document.
        for (let j = 0; j < document.Sections.get(0).Paragraphs.Count; j++) {
          let p = document.Sections.get(0).Paragraphs.get_Item(j);

          //Traverse every child object of a paragraph.
          for (let i = 0; i < p.ChildObjects.Count; i++) {
            let obj = p.ChildObjects.get(i);

            //Find the page break object.
            if (obj.DocumentObjectType == wasmModule.DocumentObjectType.Break) {
              let b = obj;

              //Remove the page break object from paragraph.
              p.ChildObjects.Remove(b);
            }
          }
        }

        // Define the output file name
        const outputFileName = "RemovePageBreaks_out.docx";

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

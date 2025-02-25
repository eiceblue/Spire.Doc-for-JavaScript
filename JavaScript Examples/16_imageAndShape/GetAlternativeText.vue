<template>
  <span>The following example demonstrates how to get the alternative text of the shape in a Word document</span>
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

        // Load the input file into the virtual file system (VFS)
        const inputFileName = "ShapeWithAlternativeText.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        //Create a document
        let document = wasmModule.Document.Create();

        //Create string builder
        let builder = [];
        document.LoadFromFile(inputFileName);

        //Loop through shapes and get the AlternativeText
        for (let i = 0; i < document.Sections.Count; i++) {
          let section = document.Sections.get(i);
          for (let j = 0; j < section.Paragraphs.Count; j++) {
            let para = section.Paragraphs.get_Item(j);
            for (let k = 0; k < para.ChildObjects.Count; k++) {
              let obj = para.ChildObjects.get(k);
              if (obj instanceof wasmModule.ShapeObject) {
                let text = obj.AlternativeText;
                //Append the alternative text in builder
                builder.push(text + "\n");
              }
            }
          }
        }

        // Convert the SB array into a Blob object
        const outputFileName = 'GetAlternativeText.txt';
        const modifiedFile = new Blob([builder.toString()], { type: "text/plain;charset=utf-8" });

        // Dispose of the document object to free resources
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
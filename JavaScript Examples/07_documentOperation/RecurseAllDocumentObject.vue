<template>
  <span>Click the following button to recurse all the document objects.</span>
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
        let inputFileName = "Sample.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();
        document.LoadFromFile(inputFileName);

        //find all document object
        let builder = [];
        for (let i = 0; i < document.Sections.Count; i++) {
          let section = document.Sections.get_Item(i);
          builder.push("section index " + i + " has following ChildObjects\n");

          for (let j = 0; j < section.Body.ChildObjects.Count; j++) {
            let obj = section.Body.ChildObjects.get(j);
            builder.push("Index : " + j + ", ChildObject Type: " + obj.DocumentObjectType + "\n");
            if (obj instanceof wasmModule.Paragraph) {
              let paragraph = obj;
              builder.push("\tParagraph index " + section.Body.GetIndex(paragraph) + " has following ChildObjects\n");
              for (let k = 0; k < paragraph.ChildObjects.Count; k++) {
                let obj2 = paragraph.ChildObjects.get(k);
                builder.push("\tIndex : " + paragraph.GetIndex(obj2) + ", ChildObject Type: " + obj2.DocumentObjectType + "\n");
              }
            }
          }
          builder.push(" \n");
        }

        // Define the output file name
        const outputFileName = "RecurseAllDocumentObject_out.txt";

        //Save to file.
        FS.writeFile(outputFileName, builder.join(""))

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: 'text/plain' });

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

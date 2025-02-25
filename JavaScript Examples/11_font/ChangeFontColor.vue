<template>
  <span>The example shows how to change Word document font color.</span>
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
        let inputFileName = "ChangeFontColor.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        
        // Create a new document
        const document = wasmModule.Document.Create();
        document.LoadFromFile(inputFileName);

        //Get the first section and first paragraph
        let section = document.Sections.get_Item(0);
        let p1 = section.Paragraphs.get_Item(0);

        //Iterate through the childObjects of the paragraph 1
        for (let i = 0; i < p1.ChildObjects.Count; i++) {
            let childObj = p1.ChildObjects.get(i);
            if (childObj instanceof wasmModule.TextRange) {
                //Change text color
                let tr = childObj;
                tr.CharacterFormat.TextColor = wasmModule.Color.get_RosyBrown();
            }
        }

        //Get the second paragraph
        let p2 = section.Paragraphs.get_Item(1);

        //Iterate through the childObjects of the paragraph 2
        for (let i = 0; i < p2.ChildObjects.Count; i++) {
            let childObj = p2.ChildObjects.get(i);
            if (childObj instanceof wasmModule.TextRange) {
                //Change text color
                let tr = childObj;
                tr.CharacterFormat.TextColor = wasmModule.Color.get_DarkGreen();
            }
        }

        // Define the output file name
        const outputFileName = "ChangeFontColor-result.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Clean up resources
        document.Dispose();

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

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

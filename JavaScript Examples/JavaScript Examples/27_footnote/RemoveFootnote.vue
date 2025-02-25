<template>
  <span>Click the following button to remove footnote in a Word document</span>
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
        let inputFileName = "Footnote.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Load a document from the virtual file system
        document.LoadFromFile(inputFileName);

        // Get the first section of the document
        let section = document.Sections.get(0);

        // Traverse paragraphs in the section to find and remove footnotes
        for (let p = 0; p < section.Paragraphs.Count; p++) {
            let para = section.Paragraphs.get_Item(p);
            let index = -1; 

            // Check each child object in the paragraph to find a footnote
            for (let i = 0, cnt = para.ChildObjects.Count; i < cnt; i++) {
                let pBase = para.ChildObjects.get(i); 
                if (pBase instanceof wasmModule.Footnote) {
                    index = i; 
                    break; 
                }
            }

            // If a footnote was found, remove it from the paragraph
            if (index > -1) {
                para.ChildObjects.RemoveAt(index); 
            }
        }
        
        // Define the output file name
        const outputFileName = "RemoveFootnote.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

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

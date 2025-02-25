<template>
  <span>Click the following button to find all hyperlinks in a Word document</span>
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
        let inputFileName = "Hyperlinks.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);        

        //Load Document
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Create a hyperlink list
        let hyperlinks = [];
        let hyperlinksText = [];
        //Iterate through the items in the sections to find all hyperlinks
        for (let i = 0; i < doc.Sections.Count; i++) {
            let section = doc.Sections.get(i);
            for (let j = 0; j < section.Body.ChildObjects.Count; j++) {
                let sec = section.Body.ChildObjects.get(j);
                if (sec.DocumentObjectType == wasmModule.DocumentObjectType.Paragraph) {
                    for (let k = 0; k < sec.ChildObjects.Count; k++) {
                        let para = sec.ChildObjects.get(k);
                        if (para.DocumentObjectType == wasmModule.DocumentObjectType.Field) {
                            let field = para;
                            if (field.Type == wasmModule.FieldType.FieldHyperlink) {
                                hyperlinks.push(field);
                                //Get the hyperlink text
                                hyperlinksText.push(field.FieldText + "\r\n");
                            }
                        }
                    }
                }
            }
        }

        // Define the output file name
        const outputFileName = "FindHyperlinks_output.txt";
        
        FS.writeFile(outputFileName,  hyperlinksText.join('\n'));

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

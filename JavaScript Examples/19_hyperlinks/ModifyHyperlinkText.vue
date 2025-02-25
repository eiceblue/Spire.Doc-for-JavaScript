<template>
  <span>Click the following button to modify hyperlink text in a Word document</span>
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

        //Find all hyperlinks in the Word document
        let hyperlinks = [];
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
                            }
                        }
                    }
                }
            }
        }

        //Reset the property of hyperlinks[0].FieldText by using the index of the hyperlink
        hyperlinks[0].FieldText = "Spire.Doc component";

        // Define the output file name
        const outputFileName = "ModifyHyperlinkText_output.docx";

        // Save the document to the specified path
        doc.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

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

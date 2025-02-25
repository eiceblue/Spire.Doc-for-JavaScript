<template>
  <span>Click the following button to remove header from Word document</span>
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
        let inputFileName = "HeaderAndFooter.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Load the document
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Get the first section of the document
        let section = doc.Sections.get_Item(0);

        //Traverse the word document and clear all headers in different type
        for (let i = 0; i < section.Paragraphs.Count; i++) {
            let para = section.Paragraphs.get_Item(i);
            for (let j = 0; j < para.ChildObjects.Count; j++) {
                //Clear footer in the first page
                let header;
                header = section.HeadersFooters.get_Item({hfType: wasmModule.HeaderFooterType.HeaderFirstPage});
                if (header != null)
                    header.ChildObjects.Clear();
                //Clear footer in the odd page
                header = section.HeadersFooters.get_Item({hfType: wasmModule.HeaderFooterType.HeaderOdd});
                if (header != null)

                    header.ChildObjects.Clear();
                //Clear footer in the even page
                header = section.HeadersFooters.get_Item({hfType: wasmModule.HeaderFooterType.HeaderEven});
                if (header != null)
                    header.ChildObjects.Clear();
            }
        }

        // Define the output file name
        const outputFileName = "RemoveHeader_output.docx";

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

<template>
  <span>Click the following button to add a header only into the first page of a Word document</span>
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
        let inputFileName = "MultiplePages.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        let inputFileName2 = "HeaderAndFooter.docx";
        await wasmModule.FetchFileToVFS(inputFileName2,"",`${import.meta.env.BASE_URL}static/data/`);

        //Load the source file
        let doc1 = wasmModule.Document.Create();
        doc1.LoadFromFile(inputFileName2);

        //Get the header from the first section
        let header = doc1.Sections.get(0).HeadersFooters.Header;

        //Load the destination file
        let doc2 = wasmModule.Document.Create();
        doc2.LoadFromFile(inputFileName);

        //Get the first page header of the destination document
        let firstPageHeader = doc2.Sections.get(0).HeadersFooters.FirstPageHeader;

        //Specify that the current section has a different header/footer for the first page
        for (let i = 0; i < doc2.Sections.Count; i++) {
            let section = doc2.Sections.get_Item(i);
            section.PageSetup.DifferentFirstPageHeaderFooter = true;
        }

        //Removes all child objects in firstPageHeader
        firstPageHeader.Paragraphs.Clear();

        //Add all child objects of the header to firstPageHeader
        for (let i = 0; i < header.ChildObjects.Count; i++) {
            let obj = header.ChildObjects.get(i);
            firstPageHeader.ChildObjects.Add(obj.Clone());
        }

        const outFileName = "AddHeaderOnlyFirstPage_output.docx";
        //Save and launch the file
        doc2.SaveToFile({fileName: outFileName,fileFormat: wasmModule.FileFormat.Docx2013});
        doc1.Close();
        doc2.Close();


        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Clean up resources
        doc1.Dispose();
        doc2.Dispose();

        // Download the file
        downloadName.value = outFileName;
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

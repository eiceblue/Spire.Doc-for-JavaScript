<template>
  <span>Click the following button to split a document into multiple documents by section break.</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from "vue";
import JSZip from "jszip";

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);

        //Define output folder 
        let outputDirectoryName = "outputFiles/";
        FS.mkdirTree(outputDirectoryName);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "Template_Docx_4.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        //Create Word document.
        let document = wasmModule.Document.Create();

        //Load the file from disk.
        document.LoadFromFile(inputFileName);

        //Define another new word document object.
        let newWord;

        //Split a Word document into multiple documents by section break.
        for (let i = 0; i < document.Sections.Count; i++) {

          newWord = wasmModule.Document.Create();
          newWord.Sections.Add(document.Sections.get(i).Clone());

          //Save to file.
          newWord.SaveToFile({ fileName: outputDirectoryName + `SplitDocBySectionBreak-${i}.docx`, fileFormat: wasmModule.FileFormat.Docx2013 });

          // Clean up resources
          newWord.Dispose();
        }

        // Clean up resources
        document.Dispose();

        const zip = new JSZip();
        let items = await FS.readdir(outputDirectoryName);
        items = items.filter((item) => item !== "." && item !== "..");
        for (const item of items) {
          const itemPath = `${outputDirectoryName}/${item}`;
          const fileData = await FS.readFile(itemPath);
          zip.file(item, fileData);
        }

        const zipBlob = await zip.generateAsync({ type: "blob" });
        const zipDownloadUrl = URL.createObjectURL(zipBlob);
        const zipDownloadName = `SplitDocBySectionBreak_out.zip`;
        downloadName.value = zipDownloadName;
        downloadUrl.value = zipDownloadUrl;
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

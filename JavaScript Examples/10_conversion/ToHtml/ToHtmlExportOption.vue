<template>
  <span>The example demonstrates how to convert Word to HTML with html export options.</span>
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
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF","/Library/Fonts/",`${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "ToHtml.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        let outputDirectoryName = "ToHTMLFolder/";
        wasmModule.FS.mkdirTree(outputDirectoryName);

        //Open a Word document.
        let document = wasmModule.Document.Create();
        document.LoadFromFile(inputFileName);

        //Set whether the css styles are embeded or not.
        document.HtmlExportOptions.CssStyleSheetFileName =  outputDirectoryName + "sample.css";
        document.HtmlExportOptions.CssStyleSheetType = wasmModule.CssStyleSheetType.External;

        //Set whether the images are embeded or not.
        document.HtmlExportOptions.ImageEmbedded = false;
        document.HtmlExportOptions.ImagesPath =  outputDirectoryName + "Demo/";

        //Set the option whether to export form fields as plain text or not.
        document.HtmlExportOptions.IsTextInputFormFieldAsText = true;

        // Define the output file name
        const outputFileName = "ToHtmlExportOption-out.html";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Html});
        // Clean up resources
        document.Dispose();

        const zip = new JSZip();
        const addFilesToZip = async (folderPath, zipFolder) => {
          let items = await FS.readdir(folderPath);
          items = items.filter((item) => item !== "." && item !== "..");
          for (const item of items) {
            const itemPath = `${folderPath}/${item}`;
            try {
              const fileData = await FS.readFile(itemPath);
              zipFolder.file(item, fileData);
            } catch (error) {
              const zipSubFolder = zipFolder.folder(item);
              await addFilesToZip(itemPath, zipSubFolder);
            }
          }
        };

        await addFilesToZip(outputDirectoryName, zip);

        const zipBlob = await zip.generateAsync({ type: "blob" });
        const zipDownloadUrl = URL.createObjectURL(zipBlob);
        const zipDownloadName = `ToHTMLFolder.zip`;
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

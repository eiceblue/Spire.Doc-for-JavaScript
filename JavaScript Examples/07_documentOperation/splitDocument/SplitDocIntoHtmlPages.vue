<template>
  <span>Click the following button to split a document into multiple html pages.</span>
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

        function IsInNextDocument(element) {
          if (element instanceof wasmModule.Paragraph) {
            let p = element;
            if (p.StyleName == "Heading1") {
              return true;
            }
          }
          return false;
        }
        function SplitDocIntoMultipleHtml(input, outDirectory) {
          //Load file
          let document = wasmModule.Document.Create();
          document.LoadFromFile(input);

          //Create a new document
          let subDoc = wasmModule.Document.Create();
          subDoc.AddSection();
          let first = true;
          let index = 0;
          for (let i = 0; i < document.Sections.Count; i++) {
            let sec = document.Sections.get_Item(i);
            for (let j = 0; j < sec.Body.ChildObjects.Count; j++) {
              let element = sec.Body.ChildObjects.get(j);
              if (IsInNextDocument(element)) {
                if (!first) {
                  //Embed css tyle and image data into html page
                  subDoc.HtmlExportOptions.CssStyleSheetType = wasmModule.CssStyleSheetType.Internal;
                  subDoc.HtmlExportOptions.ImageEmbedded = true;
                  //Save to html file
                  subDoc.SaveToFile({ fileName: outDirectory + `SplitDocIntoHtmlPages-${index}.docx`, fileFormat: wasmModule.FileFormat.Html });
                  index++;
                }
                first = false;
              }
              if (subDoc == null) {
                subDoc = wasmModule.Document.Create();
                subDoc.AddSection();
              }
              subDoc.Sections.get(0).Body.ChildObjects.Add(element.Clone());
            }
          }
          if (subDoc != null) {
            //Embed css tyle and image data into html page
            subDoc.HtmlExportOptions.CssStyleSheetType = wasmModule.CssStyleSheetType.Internal;
            subDoc.HtmlExportOptions.ImageEmbedded = true;
            //Save to html file
            subDoc.SaveToFile({ fileName: outDirectory + `SplitDocIntoHtmlPages-${index}.docx`, fileFormat: wasmModule.FileFormat.Html });
            index++;
          }
          subDoc.Close();
          document.Close();
        }

        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);

        //Define output folder 
        let outputDirectoryName = "outputFiles/";
        FS.mkdirTree(outputDirectoryName);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "SplitDocIntoHtmlPages.doc";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);


        //Split a document into multiple html pages.
        SplitDocIntoMultipleHtml(inputFileName, outputDirectoryName);

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
        const zipDownloadName = `SplitDocIntoHtmlPages_out.zip`;
        downloadName.value = zipDownloadName;
        downloadUrl.value = zipDownloadUrl;

      }
    }
    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>

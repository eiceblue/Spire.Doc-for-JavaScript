<template>
  <span>Click the following button to get the revisions of Word document.</span>
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
        let inputFileName = "GetRevisions.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        //Create a new document
        let document = wasmModule.Document.Create();

        //Load the file
        document.LoadFromFile(inputFileName);

        let insertRevision = []
        insertRevision.push("Insert revisions:\n");
        let index_insertRevision = 0;
        let deleteRevision = [];
        deleteRevision.push("Delete revisions:\n");
        let index_deleteRevision = 0;
        //Traverse sections
        for (let i = 0; i < document.Sections.Count; i++) {
          let sec = document.Sections.get_Item(i);
          //Iterate through the element under body in the section
          for (let j = 0; j < sec.Body.ChildObjects.Count; j++) {
            let docItem = sec.Body.ChildObjects.get(j);
            if (docItem instanceof wasmModule.Paragraph) {
              let para = docItem;
              //Determine if the paragraph is an insertion revision
              if (para.IsInsertRevision) {
                index_insertRevision++;
                insertRevision.push("Index: " + index_insertRevision + "\n");
                //Get insertion revision
                let insRevison = para.InsertRevision;

                //Get insertion revision type
                let insType = insRevison.Type;
                insertRevision.push("Type: " + insType + "\n");
                //Get insertion revision author
                let insAuthor = insRevison.Author;
                insertRevision.push("Author: " + insAuthor + "\n");
              }
              //Determine if the paragraph is a delete revision
              else if (para.IsDeleteRevision) {
                index_deleteRevision++;
                deleteRevision.push("Index: " + index_deleteRevision + "\n");
                let delRevison = para.DeleteRevision;
                let delType = delRevison.Type;
                deleteRevision.push("Type: " + delType + "\n");
                let delAuthor = delRevison.Author;
                deleteRevision.push("Author: " + delAuthor + "\n");
              }
              //Iterate through the element in the paragraph
              for (let i = 0; i < para.ChildObjects.Count; i++) {
                let obj = para.ChildObjects.get(i);
                if (obj instanceof wasmModule.TextRange) {
                  let textRange = obj;
                  //Determine if the textrange is an insertion revision
                  if (textRange.IsInsertRevision) {
                    index_insertRevision++;
                    insertRevision.push("Index: " + index_insertRevision + "\n");
                    let insRevison = textRange.InsertRevision;
                    let insType = insRevison.Type;
                    insertRevision.push("Type: " + insType + "\n");
                    let insAuthor = insRevison.Author;
                    insertRevision.push("Author: " + insAuthor + "\n");
                  } else if (textRange.IsDeleteRevision) {
                    index_deleteRevision++;
                    deleteRevision.push("Index: " + index_deleteRevision + "\n");
                    //Determine if the textrange is a delete revision
                    let delRevison = textRange.DeleteRevision;
                    let delType = delRevison.Type;
                    deleteRevision.push("Type: " + delType + "\n");
                    let delAuthor = delRevison.Author;
                    deleteRevision.push("Author: " + delAuthor + "\n");
                  }
                }
              }
            }
          }
        }
        // Define the output file name
        const outputFileName1 = outputDirectoryName + "insertRevision_out.txt";
        const outputFileName2 = outputDirectoryName + "deleteRevision_out.txt";
        FS.writeFile(outputFileName1, insertRevision.join(""))
        FS.writeFile(outputFileName2, deleteRevision.join(""))

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
        const zipDownloadName = `GetRevisions_out.zip`;
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

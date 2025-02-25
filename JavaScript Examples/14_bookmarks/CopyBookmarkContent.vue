<template>
  <span>The example demonstrates how to copy bookmark content in a Word document.</span>
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
        let inputFileName = "CopyBookmarkContent.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Load the document from disk.
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Get the bookmark by name.
        let bookmark = doc.Bookmarks._get_Item("Test");
        let docObj = null;

        //Judge if the paragraph includes the bookmark exists in the table, if it exists in cell,
        //Then need to find its outermost parent object(Table),
        //and get the start/end index of current object on body.
        if (bookmark.BookmarkStart.Owner.IsInCell) {
            docObj = bookmark.BookmarkStart.Owner.Owner.Owner.Owner;
        } else {
            docObj = bookmark.BookmarkStart.Owner;
        }
        let startIndex = doc.Sections.get(0).Body.ChildObjects.IndexOf(docObj);
        if (bookmark.BookmarkEnd.Owner.IsInCell) {
            docObj = bookmark.BookmarkEnd.Owner.Owner.Owner.Owner;
        } else {
            docObj = bookmark.BookmarkEnd.Owner;
        }
        let endIndex = doc.Sections.get(0).Body.ChildObjects.IndexOf(docObj);

        //Get the start/end index of the bookmark object on the paragraph.
        let para = bookmark.BookmarkStart.Owner;
        let pStartIndex = para.ChildObjects.IndexOf(bookmark.BookmarkStart);
        para = bookmark.BookmarkEnd.Owner;
        let pEndIndex = para.ChildObjects.IndexOf(bookmark.BookmarkEnd);

        //Get the content of current bookmark and copy.
        let select = wasmModule.TextBodySelection.Create(doc.Sections.get_Item(0).Body, startIndex, endIndex, pStartIndex, pEndIndex);
        let body = wasmModule.TextBodyPart.CreateByTextBodySelection(select);
        for (let i = 0; i < body.BodyItems.Count; i++) {
            doc.Sections.get(0).Body.ChildObjects.Add(body.BodyItems.get(i).Clone());
        }
        // Define the output file name
        const outputFileName = "CopyBookmarkContent-result.docx";

        // Save the document to the specified path
        doc.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Clean up resources
        doc.Dispose();

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

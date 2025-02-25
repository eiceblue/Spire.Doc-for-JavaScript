<template>
  <span>The example demonstrates how to create bookmark in a Word document.</span>
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

        //Create word document.
        let document = wasmModule.Document.Create();

        //Create a new section.
        let section = document.AddSection();

        CreateBookmarkBase(section);

        // Define the output file name
        const outputFileName = "CreateBookmark-result.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Clean up resources
        document.Dispose();

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
      function CreateBookmarkBase( section) {
        let paragraph = section.AddParagraph();
        let txtRange = paragraph.AppendText("The following example demonstrates how to create bookmark in a Word document.");
        txtRange.CharacterFormat.Italic = true;

        section.AddParagraph();
        paragraph = section.AddParagraph();
        txtRange = paragraph.AppendText("Simple Create Bookmark.");
        txtRange.CharacterFormat.TextColor = wasmModule.Color.get_CornflowerBlue();
        paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Heading2});

        //Write simple CreateBookmarks.
        section.AddParagraph();
        paragraph = section.AddParagraph();
        paragraph.AppendBookmarkStart("SimpleCreateBookmark");
        paragraph.AppendText("This is a simple bookmark.");
        paragraph.AppendBookmarkEnd("SimpleCreateBookmark");

        section.AddParagraph();
        paragraph = section.AddParagraph();
        txtRange = paragraph.AppendText("Nested Create Bookmark.");
        txtRange.CharacterFormat.TextColor = wasmModule.Color.get_CornflowerBlue();
        paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Heading2});

        //Write nested CreateBookmarks.
        section.AddParagraph();
        paragraph = section.AddParagraph();
        paragraph.AppendBookmarkStart("Root");
        txtRange = paragraph.AppendText(" This is Root data ");
        txtRange.CharacterFormat.Italic = true;
        paragraph.AppendBookmarkStart("NestedLevel1");
        txtRange = paragraph.AppendText(" This is Nested Level1 ");
        txtRange.CharacterFormat.Italic = true;
        txtRange.CharacterFormat.TextColor = wasmModule.Color.get_DarkSlateGray();
        paragraph.AppendBookmarkStart("NestedLevel2");
        txtRange = paragraph.AppendText(" This is Nested Level2 ");
        txtRange.CharacterFormat.Italic = true;
        txtRange.CharacterFormat.TextColor = wasmModule.Color.get_DimGray();
        paragraph.AppendBookmarkEnd("NestedLevel2");
        paragraph.AppendBookmarkEnd("NestedLevel1");
        paragraph.AppendBookmarkEnd("Root");
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

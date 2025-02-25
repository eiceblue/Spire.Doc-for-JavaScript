<template>
  <span>The example demonstrates how to create list styles and apply to paragraphs.</span>
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

        // Create a new document
        const document = wasmModule.Document.Create();

        //Add a section
        let sec = document.AddSection();
        //Add paragraph and set list style
        let paragraph = sec.AddParagraph();
        paragraph.AppendText("Lists");
        paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Title});

        paragraph = sec.AddParagraph();
        paragraph.AppendText("Numbered List:").CharacterFormat.Bold = true;

        //Create list style
        let numberList = wasmModule.ListStyle.Create(document, wasmModule.ListType.Numbered);
        numberList.Name = "numberList";
        numberList.Levels.get_Item(1).NumberPrefix = "%1.";
        numberList.Levels.get_Item(1).PatternType = wasmModule.ListPatternType.Arabic;
        numberList.Levels.get_Item(2).NumberPrefix = "%1.%2.";
        numberList.Levels.get_Item(2).PatternType = wasmModule.ListPatternType.Arabic;

        let bulletList = wasmModule.ListStyle.Create(document, wasmModule.ListType.Bulleted);
        bulletList.Name = "bulletList";

        //add the list style into document
        document.ListStyles.Add(numberList);
        document.ListStyles.Add(bulletList);

        //Add paragraph and apply the list style
        paragraph = sec.AddParagraph();
        paragraph.AppendText("List Item 1");
        paragraph.ListFormat.ApplyStyle(numberList.Name);

        paragraph = sec.AddParagraph();
        paragraph.AppendText("List Item 2");
        paragraph.ListFormat.ApplyStyle(numberList.Name);

        paragraph = sec.AddParagraph();
        paragraph.AppendText("List Item 2.1");
        paragraph.ListFormat.ApplyStyle(numberList.Name);
        paragraph.ListFormat.ListLevelNumber = 1;

        paragraph = sec.AddParagraph();
        paragraph.AppendText("List Item 2.2");
        paragraph.ListFormat.ApplyStyle(numberList.Name);
        paragraph.ListFormat.ListLevelNumber = 1;

        paragraph = sec.AddParagraph();
        paragraph.AppendText("List Item 2.2.1");
        paragraph.ListFormat.ApplyStyle(numberList.Name);
        paragraph.ListFormat.ListLevelNumber = 2;
        paragraph = sec.AddParagraph();
        paragraph.AppendText("List Item 2.2.2");
        paragraph.ListFormat.ApplyStyle(numberList.Name);
        paragraph.ListFormat.ListLevelNumber = 2;
        paragraph = sec.AddParagraph();
        paragraph.AppendText("List Item 2.2.3");
        paragraph.ListFormat.ApplyStyle(numberList.Name);
        paragraph.ListFormat.ListLevelNumber = 2;

        paragraph = sec.AddParagraph();
        paragraph.AppendText("List Item 2.3");
        paragraph.ListFormat.ApplyStyle(numberList.Name);
        paragraph.ListFormat.ListLevelNumber = 1;

        paragraph = sec.AddParagraph();
        paragraph.AppendText("List Item 3");
        paragraph.ListFormat.ApplyStyle(numberList.Name);

        paragraph = sec.AddParagraph();
        paragraph.AppendText("Bulleted List:").CharacterFormat.Bold = true;

        paragraph = sec.AddParagraph();
        paragraph.AppendText("List Item 1");
        paragraph.ListFormat.ApplyStyle(bulletList.Name);
        paragraph = sec.AddParagraph();
        paragraph.AppendText("List Item 2");
        paragraph.ListFormat.ApplyStyle(bulletList.Name);

        paragraph = sec.AddParagraph();
        paragraph.AppendText("List Item 2.1");
        paragraph.ListFormat.ApplyStyle(bulletList.Name);
        paragraph.ListFormat.ListLevelNumber = 1;
        paragraph = sec.AddParagraph();
        paragraph.AppendText("List Item 2.2");
        paragraph.ListFormat.ApplyStyle(bulletList.Name);
        paragraph.ListFormat.ListLevelNumber = 1;
        paragraph = sec.AddParagraph();
        paragraph.AppendText("List Item 3");
        paragraph.ListFormat.ApplyStyle(bulletList.Name);

        // Define the output file name
        const outputFileName = "Lists-result.docx";

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
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>

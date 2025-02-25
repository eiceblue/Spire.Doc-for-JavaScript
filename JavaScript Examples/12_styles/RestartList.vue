<template>
  <span>The example shows how to restart the list. </span>
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

        //Create word document
        let document = wasmModule.Document.Create();

        //Create a new section
        let section = document.AddSection();

        //Create a new paragraph
        let paragraph = section.AddParagraph();

        //Append Text
        paragraph.AppendText("List 1");

        let numberList = wasmModule.ListStyle.Create(document, wasmModule.ListType.Numbered);
        numberList.Name = "Numbered1";
        document.ListStyles.Add(numberList);

        //Add paragraph and apply the list style
        paragraph = section.AddParagraph();
        paragraph.AppendText("List Item 1");
        paragraph.ListFormat.ApplyStyle(numberList.Name);

        paragraph = section.AddParagraph();
        paragraph.AppendText("List Item 2");
        paragraph.ListFormat.ApplyStyle(numberList.Name);

        paragraph = section.AddParagraph();
        paragraph.AppendText("List Item 3");
        paragraph.ListFormat.ApplyStyle(numberList.Name);

        paragraph = section.AddParagraph();
        paragraph.AppendText("List Item 4");
        paragraph.ListFormat.ApplyStyle(numberList.Name);

        //Append Text
        paragraph = section.AddParagraph();
        paragraph.AppendText("List 2");

        let numberList2 = wasmModule.ListStyle.Create(document, wasmModule.ListType.Numbered);
        numberList2.Name = "Numbered2";
        //set start number of second list
        numberList2.Levels.get_Item(0).StartAt = 10;
        document.ListStyles.Add(numberList2);

        //Add paragraph and apply the list style
        paragraph = section.AddParagraph();
        paragraph.AppendText("List Item 5");
        paragraph.ListFormat.ApplyStyle(numberList2.Name);

        paragraph = section.AddParagraph();
        paragraph.AppendText("List Item 6");
        paragraph.ListFormat.ApplyStyle(numberList2.Name);

        paragraph = section.AddParagraph();
        paragraph.AppendText("List Item 7");
        paragraph.ListFormat.ApplyStyle(numberList2.Name);

        paragraph = section.AddParagraph();
        paragraph.AppendText("List Item 8");
        paragraph.ListFormat.ApplyStyle(numberList2.Name);


        // Define the output file name
        const outputFileName = "RestartList-result.docx";

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

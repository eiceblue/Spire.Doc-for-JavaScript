<template>
  <span>The example shows how to create bullet styles using ASCII characters.</span>
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

        //Create a new document
        let document = wasmModule.Document.Create();
        let section = document.AddSection();

        //Create four list styles based on different ASCII characters
        let listStyle1 = wasmModule.ListStyle.Create(document, wasmModule.ListType.Bulleted);
        listStyle1.Name = "liststyle";
        listStyle1.Levels.get_Item(0).BulletCharacter = String.fromCharCode(0x006e);
        listStyle1.Levels.get_Item(0).CharacterFormat.FontName = "Wingdings";
        document.ListStyles.Add(listStyle1);
        let listStyle2 = wasmModule.ListStyle.Create(document, wasmModule.ListType.Bulleted);
        listStyle2.Name = "liststyle2";
        listStyle2.Levels.get_Item(0).BulletCharacter =String.fromCharCode(0x0075);
        listStyle2.Levels.get_Item(0).CharacterFormat.FontName = "Wingdings";
        document.ListStyles.Add(listStyle2);
        let listStyle3 = wasmModule.ListStyle.Create(document, wasmModule.ListType.Bulleted);
        listStyle3.Name = "liststyle3";
        listStyle3.Levels.get_Item(0).BulletCharacter = String.fromCharCode(0x00b2);
        listStyle3.Levels.get_Item(0).CharacterFormat.FontName = "Wingdings";
        document.ListStyles.Add(listStyle3);
        let listStyle4 = wasmModule.ListStyle.Create(document, wasmModule.ListType.Bulleted);
        listStyle4.Name = "liststyle4";
        listStyle4.Levels.get_Item(0).BulletCharacter = String.fromCharCode(0x00d8);
        listStyle4.Levels.get_Item(0).CharacterFormat.FontName = "Wingdings";
        document.ListStyles.Add(listStyle4);

        //Add four paragraphs and apply list style separately
        let p1 = section.Body.AddParagraph();
        p1.AppendText("Spire.Doc for JavaScript");
        p1.ListFormat.ApplyStyle(listStyle1.Name);
        let p2 = section.Body.AddParagraph();
        p2.AppendText("Spire.Doc for JavaScript");
        p2.ListFormat.ApplyStyle(listStyle2.Name);
        let p3 = section.Body.AddParagraph();
        p3.AppendText("Spire.Doc for JavaScript");
        p3.ListFormat.ApplyStyle(listStyle3.Name);
        let p4 = section.Body.AddParagraph();
        p4.AppendText("Spire.Doc for JavaScript");
        p4.ListFormat.ApplyStyle(listStyle4.Name);

        // Define the output file name
        const outputFileName = "ASCIICharactersBulletStyle-result.docx";

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

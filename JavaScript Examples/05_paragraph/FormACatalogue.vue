<template>
  <span>The following example shows how to form a catalogue from Word headings. </span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName"> Click here to download the generated file </a>
</template>

<script>
import { ref } from 'vue';

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref('');

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);
        const convert_builtinStyleenum = ['Heading1', 'Heading2', 'Heading3'];
        //Create Word document.
        let document = wasmModule.Document.Create();

        //Add a new section.
        let section = document.AddSection();
        let paragraph = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();

        //Add Heading 1.
        paragraph = section.AddParagraph();
        paragraph.AppendText(convert_builtinStyleenum[0]);
        paragraph.ApplyStyle({
          builtinStyle: wasmModule.BuiltinStyle.Heading1,
        });
        paragraph.ListFormat.ApplyNumberedStyle();

        //Add Heading 2.
        paragraph = section.AddParagraph();
        paragraph.AppendText(convert_builtinStyleenum[1]);
        paragraph.ApplyStyle({
          builtinStyle: wasmModule.BuiltinStyle.Heading2,
        });

        //List style for Headings 2.
        let listSty2 = wasmModule.ListStyle.Create(document, wasmModule.ListType.Numbered);
        for (let i = 0; i < listSty2.Levels.Count; i++) {
          let listLev = listSty2.Levels.get_Item(i);
          listLev.UsePrevLevelPattern = true;
          listLev.NumberPrefix = '1.';
        }
        listSty2.Name = 'MyStyle2';
        document.ListStyles.Add(listSty2);
        paragraph.ListFormat.ApplyStyle(listSty2.Name);

        //Add list style 3.
        let listSty3 = wasmModule.ListStyle.Create(document, wasmModule.ListType.Numbered);
        for (let i = 0; i < listSty3.Levels.Count; i++) {
          let listLev = listSty3.Levels.get_Item(i);
          listLev.UsePrevLevelPattern = true;
          listLev.NumberPrefix = '1.1.';
        }
        listSty3.Name = 'MyStyle3';
        document.ListStyles.Add(listSty3);

        //Add Heading 3.
        for (let i = 0; i < 4; i++) {
          paragraph = section.AddParagraph();

          //Append text
          paragraph.AppendText(convert_builtinStyleenum[2]);

          //Apply list style 3 for Heading 3
          paragraph.ApplyStyle({
            builtinStyle: wasmModule.BuiltinStyle.Heading3,
          });
          paragraph.ListFormat.ApplyStyle(listSty3.Name);
        }
        // Define the output file name
        const outputFileName = 'FormACatalogue.docx';

        // Save the document to the specified path
        document.SaveToFile({
          fileName: outputFileName,
          fileFormat: wasmModule.FileFormat.Docx2013,
        });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });

        // Clean up resources
        document.Dispose();

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

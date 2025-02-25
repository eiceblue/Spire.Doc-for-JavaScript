<template>
  <span>The following example shows how to set superscript and subscript for text. </span>
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

        //Create word document
        let document = wasmModule.Document.Create();

        //Create a new section
        let section = document.AddSection();

        let paragraph = section.AddParagraph();
        paragraph.AppendText('E = mc');
        let range1 = paragraph.AppendText('2');

        //Set supperscript
        range1.CharacterFormat.SubSuperScript = wasmModule.SubSuperScript.SuperScript;
        // Insert a line break in the paragraph.
        paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
        paragraph.AppendText('F');
        let range2 = paragraph.AppendText('n');

        //Set subscript
        range2.CharacterFormat.SubSuperScript = wasmModule.SubSuperScript.SubScript;
        // Append the text " = Fn-1 + Fn-2" with specific subscripts to the paragraph.
        paragraph.AppendText(' = F');
        paragraph.AppendText('n-1').CharacterFormat.SubSuperScript = wasmModule.SubSuperScript.SubScript;
        paragraph.AppendText(' + F');
        paragraph.AppendText('n-2').CharacterFormat.SubSuperScript = wasmModule.SubSuperScript.SubScript;

        //Set font size
        for (let i = 0; i < paragraph.Items.Count; i++) {
          let range = paragraph.Items.get_Item(i);
          if (range instanceof wasmModule.TextRange) {
            range.CharacterFormat.FontSize = 36;
          }
        }
        // Define the output file name
        const outputFileName = 'SetSuperscriptAndSubscript.docx';

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

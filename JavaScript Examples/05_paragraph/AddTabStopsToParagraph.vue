<template>
  <span>The following example shows how to add tab stops to Word paragraphs. </span>
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

        //Create Word document.
        let document = wasmModule.Document.Create();

        //Add a section.
        let section = document.AddSection();

        //Add paragraph 1.
        let paragraph1 = section.AddParagraph();

        //Add tab and set its position (in points).
        let tab = paragraph1.Format.Tabs.AddTab({position: 28});

        //Set tab alignment.
        tab.Justification = wasmModule.TabJustification.Left;

        //Move to next tab and append text.
        paragraph1.AppendText('\tWashing Machine');

        //Add another tab and set its position (in points).
        tab = paragraph1.Format.Tabs.AddTab({position: 280});

        //Set tab alignment.
        tab.Justification = wasmModule.TabJustification.Left;

        //Specify tab leader type.
        tab.TabLeader = wasmModule.TabLeader.Dotted;

        //Move to next tab and append text.
        paragraph1.AppendText('\t$650');

        //Add paragraph 2.
        let paragraph2 = section.AddParagraph();

        //Add tab and set its position (in points).
        tab = paragraph2.Format.Tabs.AddTab({position: 28});

        //Set tab alignment.
        tab.Justification = wasmModule.TabJustification.Left;

        //Move to next tab and append text.
        paragraph2.AppendText('\tRefrigerator');

        //Add another tab and set its position (in points).
        tab = paragraph2.Format.Tabs.AddTab({position: 280});

        //Set tab alignment.
        tab.Justification = wasmModule.TabJustification.Left;

        //Specify tab leader type.
        tab.TabLeader = wasmModule.TabLeader.NoLeader;

        //Move to next tab and append text.
        paragraph2.AppendText('\t$800');

        // Define the output file name
        const outputFileName = 'AddTabStopsToParagraph.docx';

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

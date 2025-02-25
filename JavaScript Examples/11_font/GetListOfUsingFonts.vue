<template>
  <span>The example shows how to get the list of using fonts in the document. </span>
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
        let inputFileName = "GetListOfUsingFonts.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        let stringBuilder = [];

        let font_obj = new Map();

        //Load word document
        let document = wasmModule.Document.Create();
        document.LoadFromFile(inputFileName);

        for (let i = 0; i < document.Sections.Count; i++) {
            let section = document.Sections.get_Item(i);
            for (let j = 0; j < section.Body.Paragraphs.Count; j++) {
                let paragraph = section.Body.Paragraphs.get_Item(j);
                for (let k = 0; k < paragraph.ChildObjects.Count; k++) {
                    let obj = paragraph.ChildObjects.get(k);
                    if (obj instanceof wasmModule.TextRange) {
                        let range = obj;
                        let font = {size:  range.CharacterFormat.FontSize, name: range.CharacterFormat.FontName};
                        if (!font_obj.has(font)) {
                            font_obj.set(font, range);
                        }

                    }
                }
            }
        }
        for (let [Key,Value] of font_obj) {
          let font = Key;
          let range = Value;
          let s = "Font Name: " + font.name + ",Size: " + font.size  + ",Color: " + range.CharacterFormat.TextColor.Name;
          stringBuilder.push(s + "\n");
        }

        const outputFileName = "GetListOfUsingFonts-result.txt";
        wasmModule.FS.writeFile(outputFileName, stringBuilder.join("\n"));

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray],{type: "text/plain"});

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

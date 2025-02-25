<template>
  <span>Click the following button to split a document into multiple documents by page break.</span>
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

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "SplitWordFileByPageBreak.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        //Define output folder 
        let outputDirectoryName = "outputFolder/";
        FS.mkdirTree(outputDirectoryName);

        //Create Word document.
        let original = wasmModule.Document.Create();
        original.LoadFromFile(inputFileName);

        //Create a new word document and add a section to it.
        let newWord = wasmModule.Document.Create();
        let section = newWord.AddSection();
        original.CloneDefaultStyleTo(newWord);
        original.CloneThemesTo(newWord);
        original.CloneCompatibilityTo(newWord);

        //Split the original word document into separate documents according to page break.
        let index = 0;

        //Traverse through all sections of original document.
        for (let i = 0; i < original.Sections.Count; i++) {
          let sec = original.Sections.get(i);
          //Traverse through all body child objects of each section.
          for (let j = 0; j < sec.Body.ChildObjects.Count; j++) {
            let obj = sec.Body.ChildObjects.get(j);
            if (obj instanceof wasmModule.Paragraph) {
              let para = obj;
              sec.CloneSectionPropertiesTo(section);
              //Add paragraph object in original section into section of new document.
              section.Body.ChildObjects.Add(para.Clone());

              for (let k = 0; k < para.ChildObjects.Count; k++) {
                let parobj = para.ChildObjects.get(k);
                if (parobj instanceof wasmModule.Break && parobj.BreakType == wasmModule.BreakType.PageBreak) {
                  //Get the index of page break in paragraph.
                  let i = para.ChildObjects.IndexOf(parobj);

                  //Remove the page break from its paragraph.
                  section.Body.LastParagraph.ChildObjects.RemoveAt(i);

                  //Save the new document to a Docx file.
                  newWord.SaveToFile({ fileName: outputDirectoryName + `SplitDocByPageBreak-${index}.docx`, fileFormat: wasmModule.FileFormat.Docx2013 });
                  index++;

                  //Create a new document and add a section.
                  newWord = wasmModule.Document.Create();
                  section = newWord.AddSection();
                  original.CloneDefaultStyleTo(newWord);
                  original.CloneThemesTo(newWord);
                  original.CloneCompatibilityTo(newWord);
                  sec.CloneSectionPropertiesTo(section);
                  //Add paragraph object in original section into section of new document.
                  section.Body.ChildObjects.Add(para.Clone());
                  if (section.Paragraphs.get_Item(0).ChildObjects.Count == 0) {
                    //Remove the first blank paragraph.
                    section.Body.ChildObjects.RemoveAt(0);
                  } else {
                    //Remove the child objects before the page break.
                    while (i >= 0) {
                      section.Paragraphs.get_Item(0).ChildObjects.RemoveAt(i);
                      i--;
                    }
                  }
                }
              }
            }
            if (obj instanceof wasmModule.Table) {
              //Add table object in original section into section of new document.
              section.Body.ChildObjects.Add(obj.Clone());
            }
          }
        }

        //Save to file
        newWord.SaveToFile({ fileName: outputDirectoryName + `SplitDocByPageBreak-${index}.docx`, fileFormat: wasmModule.FileFormat.Docx2013 });
        
        // Clean up resources
        original.Dispose();
        newWord.Dispose();

        const zip = new JSZip();
        const addFilesToZip = async (folderPath, zipFolder) => {
          let items = await FS.readdir(folderPath);
          items = items.filter((item) => item !== "." && item !== "..");
          for (const item of items) {
            const itemPath = `${folderPath}/${item}`;
            try {
              const fileData = await FS.readFile(itemPath);
              zipFolder.file(item, fileData);
            } catch (error) {
              const zipSubFolder = zipFolder.folder(item);
              await addFilesToZip(itemPath, zipSubFolder);
            }
          }
        };

        await addFilesToZip(outputDirectoryName, zip);

        const zipBlob = await zip.generateAsync({ type: "blob" });
        const zipDownloadUrl = URL.createObjectURL(zipBlob);
        const zipDownloadName = `SplitDocByPageBreak_out.zip`;
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

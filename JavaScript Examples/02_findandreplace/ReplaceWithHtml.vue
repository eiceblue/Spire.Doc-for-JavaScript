<template>
  <span>The following example demonstrates how to replace content with html in a Word document. </span>
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

        // Load the sample files into the virtual file system (VFS)
        let HTMLName = 'InputHtml1.txt';
        await wasmModule.FetchFileToVFS(HTMLName, '', `${import.meta.env.BASE_URL}static/data/`);
        
        // Read html file as Uint8Array
        const htmlUintArray = wasmModule.FS.readFile(HTMLName);
        // Decode the Uint8Array to a string
        const decoder = new TextDecoder('utf-8');
        const htmlString=decoder.decode(htmlUintArray);
        
        let inputFileName = 'ReplaceWithHtml.docx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);
        //Load a template document
        let document = wasmModule.Document.Create();
        document.LoadFromFile(inputFileName);

        // Collect the objects which is used to replace text
        let replacement = [];

        // Create a temporary section
        let tempSection = document.AddSection();

        // Add a paragraph to append html
        let par = tempSection.AddParagraph();
        
        // Append the HTML content to the paragraph
        par.AppendHTML(htmlString);

        // Get the objects in temporary section
        for (let i = 0; i < tempSection.Body.ChildObjects.Count; i++) {
          let docObj = tempSection.Body.ChildObjects.get(i);
          replacement.push(docObj);
        }

        //Find all text which will be replaced.
        let selections = document.FindAllString('[#placeholder]', false, true);
        let locations = [];
        for (let selection of selections) {
          // Get the range of the current selection and create a new TextRangeLocation object with it
          locations.push(new TextRangeLocation(selection.GetAsOneRange()));
        }
        locations.sort();

        for (let location of locations) {
          //replace the text with HTML
          ReplaceWithHTML(location, replacement);
        }

        //remove the temp section
        document.Sections.Remove(tempSection);

        // Define the output file name
        const outputFileName = 'ReplaceWithHtml.docx';

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
    function ReplaceWithHTML(location, replacement) {
      let textRange = location.Text;

      //textRange index
      let index = location.Index;

      //get owener paragraph
      let paragraph = location.Owner;

      //get owner text body
      let sectionBody = paragraph.OwnerTextBody;

      //get the index of paragraph in section
      let paragraphIndex = sectionBody.ChildObjects.IndexOf(paragraph);

      let replacementIndex = -1;
      if (index === 0) {
        //remove the first child object
        paragraph.ChildObjects.RemoveAt(0);

        replacementIndex = sectionBody.ChildObjects.IndexOf(paragraph);
      } else if (index == paragraph.ChildObjects.Count - 1) {
        paragraph.ChildObjects.RemoveAt(index);
        replacementIndex = paragraphIndex + 1;
      } else {
        //split owner paragraph
        let paragraph1 = paragraph.Clone();
        while (paragraph.ChildObjects.Count > index) {
          paragraph.ChildObjects.RemoveAt(index);
        }
        let i = 0;
        let count = index + 1;
        while (i < count) {
          paragraph1.ChildObjects.RemoveAt(0);
          i += 1;
        }
        sectionBody.ChildObjects.Insert(paragraphIndex + 1, paragraph1);

        replacementIndex = paragraphIndex + 1;
      }

      //insert replacement
      for (let i = 0; i <= replacement.length - 1; i++) {
        sectionBody.ChildObjects.Insert(replacementIndex + i, replacement[i].Clone());
      }
    }
    function TextRangeLocation(TextRange) {
      this.Text = TextRange;
      this.Owner = this.Text.OwnerParagraph;
      this.Index = this.Owner.ChildObjects.IndexOf(this.Text);
      this.CompareTo = function (other) {
        return -(this.Index - other.Index);
      };
    }

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>

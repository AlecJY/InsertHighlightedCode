/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import hljs from 'highlight.js/lib/core';
import java from 'highlight.js/lib/languages/java';
import javascript from 'highlight.js/lib/languages/javascript';
import xml from 'highlight.js/lib/languages/xml';
import computedStyleToInlineStyle from 'computed-style-to-inline-style';

hljs.registerLanguage('html', xml);
hljs.registerLanguage('java', java);
hljs.registerLanguage('javascript', javascript);
hljs.registerLanguage('xml', xml);

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("insert-code").onclick = () => tryCatch(insertCode)
  }
});

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

async function insertCode() {
  await Word.run(async (context) => {
    const doc = context.document;
    const originalRange = doc.getSelection();
    const code = document.getElementById("code");
    const lang = document.getElementById("lang");
    const inTable = document.getElementById("in-table");
    const highlightedCode = lang.value === "auto"?
      hljs.highlightAuto(
        code.value
      ):
      hljs.highlight(
        code.value,
        { language: lang.value }
      );

    const result = document.getElementById("result");
    result.innerHTML = highlightedCode.value;
    computedStyleToInlineStyle(result, {
      recursive: true,
    });

    // It seems that Word doesn't support 'white-space' css
    const content = result.innerHTML.replace(/(?:\r\n|\r|\n)/g, '<br>');

    if (inTable.checked) {
      const border = 'style="border: 1px solid black;border-collapse: collapse;"';
      originalRange.insertHtml(`<table ${border}><tr ${border}><td ${border}>${content}</td></tr></table><br/>`, Word.InsertLocation.end);
    } else {
      originalRange.insertHtml(content, Word.InsertLocation.end);
    }

    await context.sync();
  })
}

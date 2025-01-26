/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import hljs from 'highlight.js/lib/core';
import go from 'highlight.js/lib/languages/go';
import java from 'highlight.js/lib/languages/java';
import javascript from 'highlight.js/lib/languages/javascript';
import kotlin from 'highlight.js/lib/languages/kotlin';
import python from 'highlight.js/lib/languages/python';
import xml from 'highlight.js/lib/languages/xml';
import computedStyleToInlineStyle from 'computed-style-to-inline-style';

hljs.registerLanguage('html', xml);
hljs.registerLanguage('go', go);
hljs.registerLanguage('java', java);
hljs.registerLanguage('javascript', javascript);
hljs.registerLanguage('kotlin', kotlin)
hljs.registerLanguage('python', python);
hljs.registerLanguage('xml', xml);

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    getThemes();
    document.getElementById("insert-code").onclick = () => tryCatch(insertCode)
  }
});

async function getThemes() {
  const themes = document.getElementsByName("theme-css");
  let currentTheme = themes[0];
  let themeList = [];

  const themeOption = document.getElementById("theme");
  themes.forEach((theme) => {
    const option = document.createElement("option");
    option.value = theme.title;
    option.innerHTML = theme.title;
    themeOption.appendChild(option);
    themeList[theme.title] = theme;
  })

  themeOption.addEventListener("change", (event) => {
    const theme = themeList[event.target.value];
    if (theme !== currentTheme) {
      theme.removeAttribute("disabled");
      currentTheme.setAttribute("disabled", "disabled");
      currentTheme = theme;
    }
  })
}

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
    const themeOption = document.getElementById("theme");
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
      let background = ''
      if (themeOption.value.endsWith('Dark')) {
        background = 'background-color: black;'
      }
      const border = `style="border: 1px solid black;border-collapse: collapse;${background}"`;
      originalRange.insertHtml(`<table class="hljs" ${border}><tr ${border}><td ${border}>${content}</td></tr></table><br/>`, Word.InsertLocation.end);
    } else {
      originalRange.insertHtml(content, Word.InsertLocation.end);
    }

    await context.sync();
  })
}

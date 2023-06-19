/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {
    const ul = document.getElementById("hits");
    ul.innerHTML = "";

    const loading = document.getElementById("loading");
    loading.style.display = "inline";

    const minWordCount = Number(document.getElementById("minWordCount").value);

    const maxWordCount = Number(document.getElementById("maxWordCount").value);

    const minOccurrences = Number(document.getElementById("minOccurrences").value);

    const maxResults = Number(document.getElementById("maxResults").value);

    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");
    await context.sync();

    // Split up the document text using existing spaces as the delimiter.
    let fulltext = paragraphs.items.reduce((a, b) => {
      return a + " " + b.text;
    }, "");

    const res = detectWordRepetitions(fulltext, minWordCount, maxWordCount, minOccurrences);
    for (let i = 0; i < res.length && i < maxResults; i++) {
      var li = document.createElement("li");
      li.onclick = function () {
        findSequence(res[i].sequence, context);
      };
      li.appendChild(document.createTextNode(res[i].sequence + " : " + res[i].hits));
      ul.appendChild(li);
    }
    loading.style.display = "none";
  });
}

function detectWordRepetitions(text, minWordCount, maxWordCount, minOccurrences) {
  const repetitions = {};

  const sentences = text.split(".");

  for (let s of sentences) {
    const words = s.trim().split(" ");
    for (let wordCount = minWordCount; wordCount <= maxWordCount; wordCount++) {
      for (let i = 0; i < words.length - wordCount; i++) {
        let sequence = words
          .slice(i, i + wordCount)
          .join(" ")
          .trim();
        const regex = new RegExp(`\\b${sequence.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, "\\$&")}\\b`, "gi");
        const matches = text.match(regex);
        // eslint-disable-next-line no-prototype-builtins
        if (matches && matches.length >= minOccurrences && !repetitions.hasOwnProperty(sequence)) {
          repetitions[sequence] = { hits: matches.length, wordCount: wordCount, sequence: sequence };
        }
      }
    }
  }

  return Object.values(repetitions).sort((a, b) => {
    if (a.wordCount == b.wordCount) {
      return b.hits - a.hits;
    } else {
      return b.wordCount - a.wordCount;
    }
  });
}

async function findSequence(sequence) {
  return Word.run(async (context) => {
    context.document.body.font.highlightColor = "white";
    const results = context.document.body.search(sequence);
    results.load("length");

    await context.sync();

    if (results.items.length > 0) {
      // Let's traverse the search results and highlight matches.
      for (let i = 0; i < results.items.length; i++) {
        results.items[i].font.highlightColor = "yellow";
      }
      results.items[0].select();
    }

    await context.sync();
  });
}

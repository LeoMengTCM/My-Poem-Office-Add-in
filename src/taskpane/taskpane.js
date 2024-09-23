/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("insert-poem").onclick = insertPoem;
    document.getElementById("analyze-poem").onclick = analyzePoem;
    document.getElementById("poet-info").onclick = poetInfo;
  }
});

export async function insertPoem() {
  return Word.run(async (context) => {
    const poem = `Do not go gentle into that good night,
Old age should burn and rave at close of day;
Rage, rage against the dying of the light.

Though wise men at their end know dark is right,
Because their words had forked no lightning they
Do not go gentle into that good night.

Good men, the last wave by, crying how bright
Their frail deeds might have danced in a green bay,
Rage, rage against the dying of the light.

Wild men who caught and sang the sun in flight,
And learn, too late, they grieved it on its way,
Do not go gentle into that good night.

Grave men, near death, who see with blinding sight
Blind eyes could blaze like meteors and be gay,
Rage, rage against the dying of the light.

And you, my father, there on the sad height,
Curse, bless, me now with your fierce tears, I pray.
Do not go gentle into that good night.
Rage, rage against the dying of the light.`;

    // Split the poem into lines
    const lines = poem.split('\n');

    // Insert each line as a separate paragraph
    for (let line of lines) {
      const paragraph = context.document.body.insertParagraph(line, Word.InsertLocation.end);
      paragraph.font.name = "Garamond";
      paragraph.font.size = 12;
      paragraph.alignment = Word.Alignment.left;
    }

    await context.sync();
  });
}

export async function analyzePoem() {
  return Word.run(async (context) => {
    const analysis = `Analysis of "Do Not Go Gentle Into That Good Night":

1. Structure: The poem is a villanelle, a 19-line poetic form with five tercets followed by a quatrain.
2. Rhyme scheme: ABA ABA ABA ABA ABA ABAA
3. Repeated lines: 
   - "Do not go gentle into that good night"
   - "Rage, rage against the dying of the light"
4. Themes:
   - Resistance to death
   - The power and limitations of human will
   - The relationship between fathers and sons
5. Imagery: Light and darkness are used as metaphors for life and death.
6. Tone: Intense, urgent, and emotional`;

    // Split the analysis into lines
    const lines = analysis.split('\n');

    // Insert each line as a separate paragraph
    for (let line of lines) {
      const paragraph = context.document.body.insertParagraph(line, Word.InsertLocation.end);
      paragraph.font.name = "Calibri";
      paragraph.font.size = 11;
      paragraph.alignment = Word.Alignment.left;
    }

    await context.sync();
  });
}

export async function poetInfo() {
  return Word.run(async (context) => {
    const info = `About Dylan Marlais Thomas:

- Born: October 27, 1914, in Swansea, Wales
- Died: November 9, 1953, in New York City, USA
- Notable works: 
  * "Do not go gentle into that good night"
  * "And death shall have no dominion"
  * "Fern Hill"
  * "Under Milk Wood" (play)
- Style: Known for his lyrical style, Thomas often employed complex imagery and sound effects in his poetry.
- Influence: Thomas's work has influenced many poets and musicians, including Bob Dylan.
- Legacy: Considered one of the most important Welsh poets of the 20th century, Thomas's work continues to be widely read and studied.`;

    // Split the info into lines
    const lines = info.split('\n');

    // Insert each line as a separate paragraph
    for (let line of lines) {
      const paragraph = context.document.body.insertParagraph(line, Word.InsertLocation.end);
      paragraph.font.name = "Arial";
      paragraph.font.size = 11;
      paragraph.alignment = Word.Alignment.left;
    }

    await context.sync();
  });
}
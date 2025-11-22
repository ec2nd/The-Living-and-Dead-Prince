# The-Living-and-Dead-Prince
NaNoGenMo 2025 Entry

A generative comparative analysis tool that transforms two translations of Machiavelli's The Prince into a single, annotated, and visually mapped PDF document.

How It WorksThe generator (main.py) ingests two raw text files (pg1232.txt and pg57037.txt) and performs a recursive structural analysis:
  Chapter Alignment: Regex pattern matching isolates chapters to ensure high-level synchronization.
  Recursive LCS Matching: A custom implementation of difflib.SequenceMatcher finds the Longest Common Subsequence.
    Matches ($\ge$ 2 words) are rendered as unified, centered text.
    Mismatches are passed recursively to find smaller matches within the divergence.
  Visual Analytics:
    Residual divergent text is subjected to NLP analysis (NLTK/VADER).
    Matplotlib generates embedded bar charts comparing Sentiment, Lexical Density, and Part-of-Speech distribution between the two translations.
  Generative Art: The system calculates the sentiment polarity of every chapter to render a unique cover image representing the emotional "heartbeat" of the novel.
  
Dependencies
  python-docx (Document generation)
  nltk (Sentiment & POS tagging)
  matplotlib (Chart generation)
  docx2pdf (Final PDF conversion)

Output
  The final artifact is a PDF titled The_Living_and_Dead_Prince.pdf, featuring a generated Table of Contents, Cover Art, and the complete dual-text comparison.

  Note: 
  pg1232.txt is a stripped version of [Niccolò Machiavelli's _The Prince_ as translated by W.K. Marriott and hosted at Project Gutenberg]([url](https://www.gutenberg.org/ebooks/1232)).
  pg57037.txt is a stripped version of [Niccolò Machiavelli's _The Prince_ as translated by Luigi Ricci and hosted at Project Gutenberg]([url](https://www.gutenberg.org/ebooks/57307)).

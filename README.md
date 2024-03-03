# Reference-Formatter 🌟
[ ] to citekey, effortlessly transforms your Word document citations into LaTeX-compatible \cite{key} references, streamlining the transition from Word to LaTeX.

## How to Use Reference-Formatter 📘

Follow these easy steps to transform your citations:

1. **Prepare Your Documents** 📄
   - **Document 1 (Word)**: Your main text with references like `[1]`, `[1, 2]`. 
   - **Document 2 (Word)**: Your reference list, e.g., `[1] paper title`,`[1] contain paper title`. 
   - **BibTeX File**: The `.bib` file with all your references in BibTeX format.

2. **Launch Reference-Formatter** 🚀
   - Open the tool and you'll see a simple, intuitive interface.

3. **Load Your Files** 📂
   - Use the `Browse` buttons to select your main document (Document 1), your reference list (Document 2), and your BibTeX file.

4. **Processing the Documents** 🔍
   - Click `Process Documents`. The tool will:
     - Find the citation number-title pairs in Document 2.
     - Match these pairs with the corresponding `citekey` in the `.bib` file.
     - Replace citation numbers in Document 1 with `\cite{citekey}`.

5. **Save Your New Document** 💾
   - Choose where to save your updated document, which is now ready for LaTeX.

6. **Check for Missing References** ⚠️
   - Reference-Formatter will notify you if any references in Document 1 are not found in Document 2 or the `.bib` file. A `missing_references.txt` file will be created for your review.

7. **Final Review** 👀
   - It's always good practice to give your document a final check to ensure all references are correctly converted and formatted.

## 🌈 Enjoy a Smoother Transition to LaTeX!

With Reference-Formatter, you can focus more on your content and less on the tedious task of reference formatting. Happy writing! 🎉

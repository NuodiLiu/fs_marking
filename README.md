
This project aims to achieve auto marking for Word/Excel exams. Currently I quickly achieved a MVP in three days. It is now used for cross checking with human markers and making sure the marking is correct and consistant. Limitation is that we have to manually change/add the rules classes if new marking critiria comes. The project is designed in a OOP way, it is highly decoupled, making it easier to add/remove new modules.

In the future, I plan to implement user customized rules:
1. User-defined rules: A DSL, close to natural language, is used to express “marking/format-checking rules.”
2. Parser: Parses the rules into an AST (Abstract Syntax Tree).
3. Selector: Works like CSS, selecting objects from the “document model” (paragraphs, headings, list levels, citations, tables, blocks, etc.).
4. Interpreter: Traverses the AST together with the selection results, calls Word COM to extract properties, and determines whether the criteria are satisfied. It then produces “rule-by-rule feedback and scores.”

To use it:
1. Make sure you are using a Windows OS, and have pywin32 lib installed in Python
2. Clone the repo
3. Make sure you have:
   1. data/submissions which contains all students MSWord final exam submission document, a reader will try to open this folder and create word instances
   2. data/output.xlsx (data folder is in the root), this is designed to align with current marking template we are using.
4. run "python3 main.py exam_feb11" or run in parallel "python main.py exam_feb11 --parallel -j 8"

Below are some example marking critiria used for a certain exam:
### ✅ Margin Checkpoints

This rule checks if the page margins are correct on the first section of the Word document.

1. **All four margins (Top, Bottom, Left, Right) must be 2.5 cm**
   - A small tolerance is allowed to avoid rounding errors.

2. **Margins are measured in points (pt)**  
   - 1 cm ≈ 28.35 pt, so 2.5 cm ≈ 70.88 pt. allowing floating points converting error.

3. **If any margin is not 2.5 cm**, it will report the actual value and give **0 mark**.

If all margins are correct → 1 mark.  
If any margin is wrong → 0 mark and show the incorrect values.

### ✅ CoverPageTitle Checkpoints

This rule checks the title on the cover page. It looks for three things:

1. **The title text should be "The Australian Ibis" (small typos are OK)**
   - It uses fuzzy match to allow small spelling mistakes.
   - The match must be at least 80% correct.

2. **The title must be a WordArt**
   - It must be a special WordArt object (not normal text with effects).
   - Only shapes on the first page are checked.

3. **The title must be roughly centered**
   - The title should be close to the center of the page (within 20pt(around 0.7cm) left or right).
   - It checks if the middle of the title is near the middle of the page.

If all 3 are OK → 1 mark.  
If any of them fail → 0 mark and shows what’s wrong.

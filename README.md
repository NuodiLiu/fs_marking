### ✅ Margin Checkpoints (Simple English)

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

### ✅ CoverPageTable Checkpoints

This rule checks the table on the cover page. It looks for:

1. **Table size must be 2 rows × 2 columns**
   - If not, it's marked as wrong.

2. **Name and ZID must be filled in correctly**
   - Cell (2,1) must have a non-empty name.
   - Cell (2,2) must contain a valid ZID (e.g., `z1234567` or just `1234567`).

3. **Table style must be correct**
   - Only these two styles are allowed:
     - "Grid Table 4 - Accent 1"
     - "Grid Table 1 - Accent 4"

4. **Table must be below the title (WordArt)**
   - The table should appear under the WordArt on the page.

If all 4 are OK → 1 mark.  
If any of them fail → 0 mark with error messages.

### ✅ TableOfContents Checkpoints

This rule checks the Table of Contents (TOC) on **page 2** of the document.

1. **There must be a Table of Contents (TOC) on page 2**
   - It checks for a TOC field (`wdFieldTOC`).
   - If no TOC is found on page 2 → **0 mark**.

2. **Page 2 must not have extra content**
   - Page 2 should only have the TOC.
   - If there is other text on page 2 (not part of TOC) → **0 mark**.

If both checks pass → **1 mark**.  
If any of them fail → **0 mark** and show an error message.

### ✅ StrictHeading1 Checkpoints (Simple English)

This rule checks if all required **Heading 1** titles are present and correct.

1. **Must use "Heading 1" style**
   - It only checks paragraphs styled as **Heading 1**.

2. **All of the following headings must be present (typos allowed):**
   - "BIOLOGY AND NATURAL HISTORY"
   - "URBAN ADAPTATION AND PUBLIC PERCEPTION"
   - "CULTURAL SIGNIFICANCE, CONSERVATION, AND FUTURE PROSPECTS"
   - "MAIN POINTS"
   - "SUMMARY"
   - "REFERENCES"
   - Small mistakes are allowed (85% fuzzy match).

3. **No extra Heading 1s**
   - If extra headings (not on the expected list) are found → error.

---

✅ **1 mark** if:
- All required headings are found (allows typos),
- No unexpected Heading 1s are present.

❌ **0 mark** if:
- Any required heading is missing, or
- There are extra headings not on the list.

### ✅ StrictHeading2 Checkpoints

This rule checks if all required **Heading 2** titles are present and correct.

1. **Must use "Heading 2" style**
   - Only paragraphs with **Heading 2** style are checked.

2. **All of the following headings must be present (typos allowed):**
   - "Physical Characteristics"
   - "Diet and Foraging Behaviour"
   - "Breeding and Life Cycle"
   - "From Wetlands to Concrete Jungles"
   - "Survivors of the Anthropocene"
   - "Public Image and Media Representation"
   - "Cultural and Indigenous Significance"
   - "Conservation Status and Concerns"
   - "Coexistence and the Future"
   - Minor typos are okay (uses 85% fuzzy match).

3. **No extra Heading 2s**
   - If any unexpected Heading 2 titles are found, it’s an error.

---

✅ **1 mark** if:
- All required headings are found (even with small typos),
- No unexpected Heading 2s are present.

❌ **0 mark** if:
- Any required heading is missing, or
- There are extra headings not on the list.


### ✅ CombinedStyle Checkpoints

This rule checks **style settings** for Normal, Heading 1, and Heading 2.  
It includes **11 style rules** in total and gives up to **2 marks** based on how many are correct.

---

#### 🔹 Normal style must:
1. Use **Arial** font  
2. Use **11pt** font size  
3. Have **single line spacing**  
4. Be **left aligned**

---

#### 🔹 Heading 1 must:
5. Use **20pt** font size  
6. Be **left aligned**  
7. Use a **blue-ish font color**  
8. Have **12pt spacing before** and **6pt after**

---

#### 🔹 Heading 2 must:
9. Be **left aligned**  
10. Be **bold**  
11. Be **italic**

---

### ✅ Marking:
- **2 marks**: All 11 checks passed  
- **1 mark**: At least 6 passed  
- **0 marks**: Fewer than 6 passed or an error occurred

Any failed checks will be listed as error messages.


### ✅ ImageRightOfText Checkpoints

This rule checks that a picture is correctly placed next to the paragraph beginning with:

> "The Australian White Ibis (Threskiornis molucca) is a distinctive bird species..."

It checks the following:

1. **The picture must be anchored to the correct paragraph**
   - Uses fuzzy match to find the correct paragraph.
   - If no matching paragraph is found → 0 mark.

2. **The picture must use “Tight” text wrapping**
   - If not set to tight wrap → 0 mark.

3. **The picture must be on the right side of the page**
   - Its center must be to the right of the page center.

4. **The picture must not be too wide**
   - Width must be less than two-thirds of the page width in case too big, it will look weird.

5. **The picture must not be too tall**
   - Height must not be greater than its width, keep orginal ratio, making sure it's not resized weirdly.

---

✅ **1 mark** if all above are correct.  
❌ **0 mark** if the image is missing or any condition fails.


### ✅ FootnoteOnHabitat Checkpoints

This rule checks if the word **"habitat"** has a correct footnote.

1. **Find the first appearance of the word "habitat"**
   - It must be somewhere in the document.

2. **The word "habitat" must have a footnote**
   - If no footnote → 0 mark.

3. **The footnote text must explain "Natural home or environment"**
   - Small wording differences are allowed (75% fuzzy match).
   - If the meaning is wrong or missing → 0 mark.

---

✅ **1 mark** if:
- "habitat" is found,
- it has a footnote,
- the footnote is roughly correct.

❌ **0 mark** if:
- "habitat" is missing,
- or it has no footnote,
- or the footnote content is incorrect.


### ✅ Footer Checkpoints

This rule checks if the document footer is correctly set up.
This checking is strict. Any footer content error will be marked to review.

---

### ✅ If the document has **only one section**:
1. **Footer must exist**  
2. **Footer must include all three parts:**
   - Left: Name  
   - Center: Page number (auto field)  
   - Right: ZID  
3. ✅ If all correct → **1 mark**  
   ❌ If missing or incorrect → **0 mark**

---

### ✅ If the document has **two sections** (most common case):
1. **Section 1 must have no footer**  
2. **Section 2 must have a valid footer**, with:
   - Left: Name  
   - Center: Page number (auto field)  
   - Right: ZID  
   - Page number must be inserted using a **field**, not typed manually  
3. ✅ If all correct → **2 marks**  
   ⚠️ If footer exists but has issues → **1 mark**  
   ❌ If footer missing or invalid → **0 mark**

---

### ❌ Errors may include:
- Missing footer in section 2  
- Footer text not split into three parts  
- No page number field detected  
- Section 1 has a footer when it shouldn’t

All problems are listed with messages, and partial marks are given when appropriate.


### ✅ Multilevel List under Main Points

This rule checks if the list under the **“Main Points”** heading is a correctly formatted multilevel list.  
It searches for 11 expected lines (like “Appearance”, “Habitat”, “Diet” and subpoints) using fuzzy match, then runs **5 checks** on formatting.

---

### 📋 The list must:

1. ✅ Be a **Word multilevel list** (not manually typed)
2. ✅ Use **A/B/C** for Level 1 bullet style  
3. ✅ Level 1 should be:
   - Aligned at **0 cm**
   - First line indent **1 cm**
4. ✅ Level 2 bullet must **not be a solid circle**  
5. ✅ Level 2 should be:
   - Aligned at **1 cm**
   - First line indent **2 cm**

---

### 🧠 Marking:

- **1 mark**: if **3 or more** of the above checks pass  
- **0 mark**: if **fewer than 3** pass

Each failed check will appear with a description of what went wrong.



### ✅ Page Break Before Heading 'SUMMARY'

This rule checks that the heading **“SUMMARY”** starts on a new page using a proper page break.

---

### ✅ The check:

1. Finds the **last appearance** with the text `"SUMMARY"` (case-insensitive).
2. Checks if the paragraph **before it** ends with a **manual page break** (`\x0c`).
3. Ignores TOC entries by searching from the end of the document.

---

### 🧠 Marking:

- ✅ **1 mark** if there is a correct page break before "SUMMARY".
- ❌ **0 mark** if:
  - "SUMMARY" is missing,
  - It is the first paragraph,
  - Or no page break is found immediately before it.

Any problem is reported with an error message.



### ✅ Hanging Indent for References

This rule checks that the 3 specific references at the end of the document:

1. Are present in the correct order (fuzzy match allowed),
2. Use a **2 cm hanging indent** formatting.

---

### ✅ The rule looks for these references:

1. *Australian White Ibis - BirdLife Australia. Retrieved from ...*
2. *Urban Pest or Aussie Hero? Changing Media Representations ...*
3. *Australian White Ibis - The Australian Museum. Retrieved from ...*

Minor typos or formatting differences are allowed (85% similarity).

---

### 🧠 Marking:

- ✅ **1 mark** if all 3 references are found and each has the correct 2cm **hanging indent**.
- ❌ **0 mark** if:
  - Fewer than 3 references are found,
  - Or any of the matched references does not use proper hanging indent.

Any issue will be reported with specific error messages (e.g. which line failed).

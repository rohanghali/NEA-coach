# Put your mark scheme in this folder

Anything in here is read and used as the criteria the AI judges your work against.
Delete the example files once you have added the real thing.

**Accepted file types:** `.md`, `.txt`, `.json`

**Not read yet:** `.pdf` — the app will tell you it skipped a PDF rather than
quietly ignoring it. Convert the pages you need to text, or ask for PDF support.

---

## Best option: a `.json` file

JSON lets criteria be named and weighted, which produces the clearest coverage
grid. Copy `example-spec.json` in this folder, replace it with your real criteria,
and delete the original.

```json
{
  "criteria": [
    {
      "id": "c1",
      "name": "Client and problem definition",
      "weight": 10,
      "description": "Identifies a real client and a justified design problem.",
      "evidence": "A clear brief, client details, and a defined need."
    }
  ]
}
```

`weight` is optional. If your mark scheme gives marks or percentages, put them
here and the coverage grid will show how much of the total you are evidencing.

---

## Also fine: plain text or Markdown

Paste the specification in as normal text. The app will pick up headings as
criteria, so use clear headings:

```markdown
## Research and analysis (20 marks)

Focused research into existing products, materials and users, clearly linked to
the design decisions that follow.
```

Markdown headings become criteria automatically. Plain text works too, but with
headings you get a much better coverage grid.

---

## Adding the real Eduqas DTA specification

1. Open the specification PDF.
2. Select the assessment criteria and mark scheme pages.
3. Copy the text into a new `.md` file here, keeping the criterion headings intact.
4. In the app, press **"Re-check everything"**.

The advice gets noticeably more specific once the real criteria are loaded,
because every suggestion can then be tied to a named criterion.

Until then, the AI falls back to general UK A-level D&T NEA expectations, and
says so in your report rather than pretending to know your mark scheme.

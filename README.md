# Mail Merge (Node.js + Carbone)

Generate personalized Word documents from a template and CSV data.

## Setup

```bash
npm install
```

## Usage

1. Create a `.docx` template with placeholders like `{d.name}`, `{d.email}`, `{d.company}`
2. Prepare your CSV with column headers matching the placeholders
3. Run:

```bash
node index.js template.docx sample_data.csv
```

Output files will be saved in the `output/` folder.

## Template Example

In your Word document, use placeholders like:

```
Dear {d.name},

Thank you for your payment of ${d.amount} from {d.company}.
```

**Note:** Carbone uses `{d.fieldName}` syntax where `d` refers to the data object.

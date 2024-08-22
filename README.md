# qexport-xlsx

async toXLSX(table: Array<Array<string>>, outputPath: string);

## Installation

```bash
npm install qexport-xlsx
```

## Usage

### Example:

```typescript
import { toXLSX } from 'qexport-xlsx'
// let { toXLSX } = require('qexport-xlsx');

toXLSX([
  ["A", "B"],
  ["C", "D"]
], "./output.xlsx").then(()=>{
  console.log("Done!");
});
```

## License

BSD 2-Clause License

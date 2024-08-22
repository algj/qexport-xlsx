import { toXLSX } from './index.js';

toXLSX([
  ["A", "B"],
  ["C", "D"]
], "./output.xlsx").then(()=>{
  console.log("Done!");
});

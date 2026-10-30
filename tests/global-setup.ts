import { resolveEmbeddedBinary } from '../packages/ooxml-validator/src/resolve-binary.js';

export default function setup() {
  try {
    resolveEmbeddedBinary();
  } catch {
    throw new Error(
      'OOXML validator binary not found. Run: npm run build:native --workspace=@attentionform/ooxml-validator (needs the .NET 10 SDK)',
    );
  }
}

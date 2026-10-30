import { describe, it, expect } from 'vitest';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { DocumentAST } from '@attentionform/astra-engine';
import { executeTable as executeTableEngine, TableOperation } from '../../packages/engine/src/operations/table';
import { executeTable as executeTableBackend } from '@attentionform/astra-engine';

function createMockMultiTableAST(): { ast: DocumentAST; initialXml1: string; initialXml2: string } {
  const xml1 = `
    <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:tblPr/>
      <w:tblGrid>
        <w:gridCol w:w="3000"/>
        <w:gridCol w:w="3000"/>
      </w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Header 1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Header 2</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Cell 1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Cell 2</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  `.trim();

  const xml2 = `
    <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:tblPr/>
      <w:tblGrid>
        <w:gridCol w:w="2000"/>
      </w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Table 2 Cell</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  `.trim();

  const doc1 = new DOMParser().parseFromString(xml1, 'text/xml');
  const doc2 = new DOMParser().parseFromString(xml2, 'text/xml');

  const ast: DocumentAST = {
    metadata: {},
    sections: [{
      id: 'sec_root',
      heading: null,
      nodes: [
        {
          id: 't_001',
          type: 'table',
          xml: doc1.documentElement,
          metadata: {
            text: 'Header 1 Header 2 Cell 1 Cell 2',
            tableDimensions: { rows: 2, cols: 2 }
          } as any,
          version: 1,
          checksum: 'chk_tbl_1'
        },
        {
          id: 't_002',
          type: 'table',
          xml: doc2.documentElement,
          metadata: {
            text: 'Table 2 Cell',
            tableDimensions: { rows: 1, cols: 1 }
          } as any,
          version: 1,
          checksum: 'chk_tbl_2'
        }
      ],
      pageSetup: {},
      version: 1,
      checksum: 'sec_chk_1'
    }],
    styles: {},
    numbering: { abstractNum: {}, concreteNum: {} },
    relationships: {},
    headers: {},
    footers: {},
    mediaFiles: {},
    settings: { trackChanges: false, defaultAuthor: 'ASTra' },
    nextChangeId: 1
  };

  const serializer = new XMLSerializer();
  return {
    ast,
    initialXml1: serializer.serializeToString(ast.sections[0].nodes[0].xml),
    initialXml2: serializer.serializeToString(ast.sections[0].nodes[1].xml),
  };
}

function createEmptyDocAST(): DocumentAST {
  return {
    metadata: {},
    sections: [{
      id: 'sec_root',
      heading: null,
      nodes: [
        {
          id: 'p_001',
          type: 'paragraph',
          xml: new DOMParser().parseFromString('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Hello</w:t></w:r></w:p>', 'text/xml').documentElement,
          metadata: { text: 'Hello' } as any,
          version: 1,
          checksum: 'chk_p_1'
        }
      ],
      pageSetup: {},
      version: 1,
      checksum: 'sec_chk_1'
    }],
    styles: {},
    numbering: { abstractNum: {}, concreteNum: {} },
    relationships: {},
    headers: {},
    footers: {},
    mediaFiles: {},
    settings: { trackChanges: false, defaultAuthor: 'ASTra' },
    nextChangeId: 1
  };
}

describe.each([
  { name: 'packages/engine executeTable', executeTable: executeTableEngine },
  { name: 'backend executeTable', executeTable: executeTableBackend },
])('Table-Ghost Target Resolution: $name', ({ executeTable }) => {
  const serializer = new XMLSerializer();

  it('throws an error containing "Table target \'table_nonexistent\' not found" and does NOT mutate first table (op: updateCell)', async () => {
    const { ast, initialXml1, initialXml2 } = createMockMultiTableAST();
    const firstTable = ast.sections[0].nodes[0];
    const initialVersion = firstTable.version;
    const initialChecksum = firstTable.checksum;
    const initialText = firstTable.metadata.text;

    const op: TableOperation = {
      action: 'table',
      target: 'table_nonexistent',
      params: {
        op: 'updateCell',
        rowIndex: 1,
        colIndex: 1,
        text: 'GHOST_MUTATION'
      }
    };

    await expect(executeTable(op, ast)).rejects.toThrowError(
      /Table target 'table_nonexistent' not found/
    );

    // Verify error lists available table IDs
    try {
      await executeTable(op, ast);
      expect.unreachable('Should have thrown');
    } catch (err: any) {
      expect(err.message).toContain("Table target 'table_nonexistent' not found");
      expect(err.message).toContain('Available table nodes: t_001, t_002');
      expect(err.message).toContain("Expected shape: { action:'table', target:'<tableId>'");
    }

    // Verify first table is NOT mutated (byte/content identical)
    expect(serializer.serializeToString(firstTable.xml)).toBe(initialXml1);
    expect(firstTable.version).toBe(initialVersion);
    expect(firstTable.checksum).toBe(initialChecksum);
    expect(firstTable.metadata.text).toBe(initialText);

    // Verify second table is also untouched
    const secondTable = ast.sections[0].nodes[1];
    expect(serializer.serializeToString(secondTable.xml)).toBe(initialXml2);
  });

  it('throws an error when flat alias operation: "addRow" is used with nonexistent target and does NOT mutate first table', async () => {
    const { ast, initialXml1 } = createMockMultiTableAST();
    const firstTable = ast.sections[0].nodes[0];
    const initialVersion = firstTable.version;
    const initialChecksum = firstTable.checksum;

    const op: TableOperation = {
      action: 'table',
      target: 'table_nonexistent',
      operation: 'addRow',
      text: 'row'
    } as any;

    await expect(executeTable(op, ast)).rejects.toThrowError(
      /Table target 'table_nonexistent' not found/
    );

    try {
      await executeTable(op, ast);
      expect.unreachable('Should have thrown');
    } catch (err: any) {
      expect(err.message).toContain("Table target 'table_nonexistent' not found");
      expect(err.message).toContain('Available table nodes: t_001, t_002');
    }

    // Verify first table is NOT mutated
    expect(serializer.serializeToString(firstTable.xml)).toBe(initialXml1);
    expect(firstTable.version).toBe(initialVersion);
    expect(firstTable.checksum).toBe(initialChecksum);
  });

  it('throws an error when tableId alias is nonexistent and does NOT mutate first table', async () => {
    const { ast, initialXml1 } = createMockMultiTableAST();
    const firstTable = ast.sections[0].nodes[0];

    const op: TableOperation = {
      action: 'table',
      tableId: 'table_nonexistent',
      params: {
        op: 'updateCell',
        rowIndex: 1,
        colIndex: 2,
        text: 'MUTATE_VIA_TABLEID'
      }
    };

    await expect(executeTable(op, ast)).rejects.toThrowError(
      /Table target 'table_nonexistent' not found/
    );

    expect(serializer.serializeToString(firstTable.xml)).toBe(initialXml1);
  });

  it('lists "Available table nodes: none" when document has no tables and nonexistent target is given', async () => {
    const ast = createEmptyDocAST();

    const op: TableOperation = {
      action: 'table',
      target: 'table_nonexistent',
      params: {
        op: 'updateCell',
        rowIndex: 1,
        colIndex: 1,
        text: 'fail'
      }
    };

    await expect(executeTable(op, ast)).rejects.toThrowError(
      /Table target 'table_nonexistent' not found\. Available table nodes: none\./
    );
  });

  it('omitting target entirely continues to locate the first table as backward-compatible fallback', async () => {
    const { ast, initialXml2 } = createMockMultiTableAST();
    const firstTable = ast.sections[0].nodes[0];

    const op: TableOperation = {
      action: 'table',
      params: {
        op: 'updateCell',
        rowIndex: 1,
        colIndex: 1,
        text: 'FALLBACK_HEADER'
      }
    };

    await executeTable(op, ast);

    // Verify first table was updated
    expect(firstTable.metadata.text).toContain('FALLBACK_HEADER');
    const firstTableXml = serializer.serializeToString(firstTable.xml);
    expect(firstTableXml).toContain('FALLBACK_HEADER');

    // Verify second table was NOT touched
    const secondTable = ast.sections[0].nodes[1];
    expect(serializer.serializeToString(secondTable.xml)).toBe(initialXml2);
  });
});

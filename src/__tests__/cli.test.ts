/**
 * @file cli.test.ts
 * Unit tests for CLI parser and runner.
 */

import { randomUUID } from 'node:crypto';
import { readFile, unlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import { getVersion, parseCliArgs, printHelp, runCli } from '../cli.js';

describe('CLI module', () => {
  const tempFiles: string[] = [];
  let logSpy: jest.SpyInstance;
  let errorSpy: jest.SpyInstance;

  async function createTempDocx(text: string = 'Test Content'): Promise<string> {
    const zip = new JSZip();
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>${text}</w:t></w:r></w:p>
  </w:body>
</w:document>`;
    zip.file('word/document.xml', docXml);
    const buffer = await zip.generateAsync({ type: 'nodebuffer' });
    const tempPath = path.join(tmpdir(), `cli-test-${randomUUID()}.docx`);
    await writeFile(tempPath, buffer);
    tempFiles.push(tempPath);
    return tempPath;
  }

  beforeEach(() => {
    logSpy = jest.spyOn(console, 'log').mockImplementation(() => {});
    errorSpy = jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(async () => {
    logSpy.mockRestore();
    errorSpy.mockRestore();

    while (tempFiles.length > 0) {
      const file = tempFiles.pop();
      if (file) {
        try {
          await unlink(file);
        } catch {
          // ignore
        }
      }
    }
  });

  describe('parseCliArgs', () => {
    it('should parse defaults when only file arguments are provided', () => {
      const opts = parseCliArgs(['file1.docx', 'file2.docx']);
      expect(opts.files).toEqual(['file1.docx', 'file2.docx']);
      expect(opts.textOnly).toBe(false);
      expect(opts.stdout).toBe(false);
      expect(opts.header).toBe(false);
      expect(opts.imagesDir).toBe('images');
      expect(opts.noEscape).toBe(false);
      expect(opts.help).toBe(false);
      expect(opts.version).toBe(false);
    });

    it('should parse --help and -h flags', () => {
      expect(parseCliArgs(['--help']).help).toBe(true);
      expect(parseCliArgs(['-h']).help).toBe(true);
    });

    it('should parse --version and -v flags', () => {
      expect(parseCliArgs(['--version']).version).toBe(true);
      expect(parseCliArgs(['-v']).version).toBe(true);
    });

    it('should parse conversion mode flags', () => {
      const opts = parseCliArgs(['--text-only', '--stdout', '--header', '--no-escape']);
      expect(opts.textOnly).toBe(true);
      expect(opts.stdout).toBe(true);
      expect(opts.header).toBe(true);
      expect(opts.noEscape).toBe(true);
    });

    it('should parse --images-dir with separate argument', () => {
      const opts = parseCliArgs(['--images-dir', 'custom_assets', 'doc.docx']);
      expect(opts.imagesDir).toBe('custom_assets');
      expect(opts.files).toEqual(['doc.docx']);
    });

    it('should parse --images-dir=value syntax', () => {
      const opts = parseCliArgs(['--images-dir=custom_assets', 'doc.docx']);
      expect(opts.imagesDir).toBe('custom_assets');
      expect(opts.files).toEqual(['doc.docx']);
    });

    it('should throw an error when --images-dir has no following argument', () => {
      expect(() => parseCliArgs(['--images-dir'])).toThrow(
        'Option "--images-dir" requires a directory path argument.',
      );
    });

    it('should throw an error when --images-dir is followed by another flag', () => {
      expect(() => parseCliArgs(['--images-dir', '--stdout'])).toThrow(
        'Option "--images-dir" requires a directory path argument.',
      );
    });

    it('should throw an error when --images-dir= is empty', () => {
      expect(() => parseCliArgs(['--images-dir='])).toThrow(
        'Option "--images-dir" requires a directory path argument.',
      );
    });

    it('should warn on unrecognized options', () => {
      const opts = parseCliArgs(['--unknown-flag', 'doc.docx']);
      expect(errorSpy).toHaveBeenCalledWith('Warning: Unrecognized option "--unknown-flag".');
      expect(opts.files).toEqual(['doc.docx']);
    });
  });

  describe('printHelp and getVersion', () => {
    it('should print usage information in printHelp', () => {
      printHelp();
      expect(logSpy).toHaveBeenCalled();
      const output = logSpy.mock.calls[0][0];
      expect(output).toContain('Usage: docx2typst');
      expect(output).toContain('--text-only');
      expect(output).toContain('--stdout');
      expect(output).toContain('--header');
      expect(output).toContain('--images-dir');
      expect(output).toContain('--no-escape');
    });

    it('should return package version from getVersion', async () => {
      const version = await getVersion();
      expect(version).toBe('0.1.0');
    });
  });

  describe('runCli execution', () => {
    it('should handle --help and exit with code 0', async () => {
      const code = await runCli(['--help']);
      expect(code).toBe(0);
      expect(logSpy).toHaveBeenCalled();
      expect(logSpy.mock.calls[0][0]).toContain('Usage: docx2typst');
    });

    it('should handle -v and exit with code 0', async () => {
      const code = await runCli(['-v']);
      expect(code).toBe(0);
      expect(logSpy).toHaveBeenCalledWith('0.1.0');
    });

    it('should print error and return 1 on invalid argument', async () => {
      const code = await runCli(['--images-dir']);
      expect(code).toBe(1);
      expect(errorSpy).toHaveBeenCalledWith(
        'Error: Option "--images-dir" requires a directory path argument.',
      );
    });

    it('should print usage and "No files converted." when no input files are provided', async () => {
      const code = await runCli([]);
      expect(code).toBe(1);
      expect(logSpy).toHaveBeenCalledWith('No files converted.');
    });

    it('should convert a docx file and write .typ file next to it', async () => {
      const docxPath = await createTempDocx('Hello Typst');
      const expectedTypPath = path.join(
        path.dirname(docxPath),
        `${path.basename(docxPath, '.docx')}.typ`,
      );
      tempFiles.push(expectedTypPath);

      const code = await runCli([docxPath]);
      expect(code).toBe(0);
      expect(logSpy).toHaveBeenCalledWith('Converted 1 file(s) successfully.');

      const writtenContent = await readFile(expectedTypPath, 'utf-8');
      expect(writtenContent).toContain('Hello Typst');
    });

    it('should output to stdout when --stdout is specified', async () => {
      const docxPath = await createTempDocx('Stdout Content');
      const code = await runCli(['--stdout', docxPath]);
      expect(code).toBe(0);

      const stdoutCalls = logSpy.mock.calls.map((call) => call[0]);
      expect(stdoutCalls.some((call) => typeof call === 'string' && call.includes('Stdout Content'))).toBe(true);
      expect(stdoutCalls).toContain('Converted 1 file(s) successfully.');
    });

    it('should extract plain text when --text-only is specified', async () => {
      const docxPath = await createTempDocx('Plain Text Only');
      const expectedTypPath = path.join(
        path.dirname(docxPath),
        `${path.basename(docxPath, '.docx')}.typ`,
      );
      tempFiles.push(expectedTypPath);

      const code = await runCli(['--text-only', docxPath]);
      expect(code).toBe(0);

      const writtenContent = await readFile(expectedTypPath, 'utf-8');
      expect(writtenContent).toBe('Plain Text Only');
    });

    it('should include header comment when --header is specified', async () => {
      const docxPath = await createTempDocx('Header Test');
      const expectedTypPath = path.join(
        path.dirname(docxPath),
        `${path.basename(docxPath, '.docx')}.typ`,
      );
      tempFiles.push(expectedTypPath);

      const code = await runCli(['--header', docxPath]);
      expect(code).toBe(0);

      const writtenContent = await readFile(expectedTypPath, 'utf-8');
      expect(writtenContent.startsWith('// Generated by docx2typst')).toBe(true);
    });

    it('should include header with --text-only and --header', async () => {
      const docxPath = await createTempDocx('Text Header Test');
      const expectedTypPath = path.join(
        path.dirname(docxPath),
        `${path.basename(docxPath, '.docx')}.typ`,
      );
      tempFiles.push(expectedTypPath);

      const code = await runCli(['--text-only', '--header', docxPath]);
      expect(code).toBe(0);

      const writtenContent = await readFile(expectedTypPath, 'utf-8');
      expect(writtenContent).toContain('// Generated by docx2typst');
      expect(writtenContent).toContain('Text Header Test');
    });

    it('should not escape special characters when --no-escape is specified', async () => {
      const docxPath = await createTempDocx('Special: @test #hash &lt;tag&gt;');
      const expectedTypPath = path.join(
        path.dirname(docxPath),
        `${path.basename(docxPath, '.docx')}.typ`,
      );
      tempFiles.push(expectedTypPath);

      const code = await runCli(['--no-escape', docxPath]);
      expect(code).toBe(0);

      const writtenContent = await readFile(expectedTypPath, 'utf-8');
      expect(writtenContent).toContain('@test');
      expect(writtenContent).not.toContain('\\@test');
    });

    it('should handle nonexistent files gracefully and return code 1', async () => {
      const nonexistentPath = path.join(tmpdir(), 'nonexistent-docx2typst.docx');
      const code = await runCli([nonexistentPath]);
      expect(code).toBe(1);
      expect(errorSpy).toHaveBeenCalled();
      expect(logSpy).toHaveBeenCalledWith('No files converted.');
    });

    it('should continue to next file when one file fails in a batch', async () => {
      const validDocx = await createTempDocx('Valid Batch Item');
      const invalidPath = path.join(tmpdir(), 'missing-batch-item.docx');
      const expectedTypPath = path.join(
        path.dirname(validDocx),
        `${path.basename(validDocx, '.docx')}.typ`,
      );
      tempFiles.push(expectedTypPath);

      const code = await runCli([invalidPath, validDocx]);
      expect(code).toBe(0);
      expect(errorSpy).toHaveBeenCalled();
      expect(logSpy).toHaveBeenCalledWith('Converted 1 file(s) successfully.');

      const writtenContent = await readFile(expectedTypPath, 'utf-8');
      expect(writtenContent).toContain('Valid Batch Item');
    });
  });
});

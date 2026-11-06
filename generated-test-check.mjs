import path from 'node:path';
import { createRequire } from 'node:module';

const MAX_MESSAGES = 20;
const MAX_MESSAGE_LENGTH = 700;
const SCOPE_WARNING = 'Static generated-source check only; not a full repo typecheck. Unrelated preexisting source/dependency diagnostics are not blocking or exhaustively collected. This does not prove scenario coverage, assertion quality, test discovery, or runtime success.';

/** Check a draft without writing it, emitting output, or executing project code. */
export async function checkGeneratedTest(source, { repositoryPath, filePath } = {}) {
  const issues = [];
  const warnings = [SCOPE_WARNING];
  const add = (messages, message) => {
    const text = String(message).replace(/\s+/g, ' ').trim().slice(0, MAX_MESSAGE_LENGTH);
    if (messages.includes(text)) return;
    if (messages.length < MAX_MESSAGES) messages.push(text);
    else messages[MAX_MESSAGES - 1] = 'Additional messages omitted (limit reached).';
  };
  const result = () => ({
    issues,
    warnings,
    summary: issues.length
      ? 'Generated test failed static checks; resolve the reported issues before use.'
      : 'Generated test passed limited static checks; review coverage and runtime prerequisites before execution.',
  });
  if (typeof source !== 'string' || !source.trim()
      || typeof repositoryPath !== 'string' || !repositoryPath
      || typeof filePath !== 'string' || !filePath) {
    add(issues, 'Nonempty source, repositoryPath, and intended final filePath are required.');
    return result();
  }

  const repository = path.resolve(repositoryPath);
  const intendedPath = path.resolve(repository, filePath);
  let ts;
  try {
    // Only the installed compiler is loaded. No project imports/config JS run.
    ts = createRequire(path.join(repository, 'package.json'))('typescript');
  } catch {
    add(issues, 'Installed TypeScript is unavailable from repository/package.json; validation cannot proceed.');
    return result();
  }

  try {
    const canonical = (name) => {
      const absolute = path.resolve(name);
      return ts.sys.useCaseSensitiveFileNames ? absolute : absolute.toLowerCase();
    };
    const isVirtual = (name) => canonical(name) === canonical(intendedPath);
    const diagnostic = (entry) => {
      let location = '';
      if (entry.file && entry.start !== undefined) {
        const { line, character } = entry.file.getLineAndCharacterOfPosition(entry.start);
        location = `${path.relative(repository, entry.file.fileName)}:${line + 1}:${character + 1}: `;
      }
      add(entry.category === ts.DiagnosticCategory.Error ? issues : warnings,
        `${location}TS${entry.code}: ${ts.flattenDiagnosticMessageText(entry.messageText, ' ')}`);
    };
    // The config parser expects normalized separators even on Windows; mixed
    // separators can trigger an internal assertion when reporting syntax errors.
    const configPath = path.join(repository, 'tsconfig.json').replaceAll('\\', '/');
    if (!ts.sys.fileExists(configPath)) {
      add(issues, 'Repository tsconfig.json is unavailable; validation cannot proceed.');
      return result();
    }
    const parsed = ts.getParsedCommandLineOfConfigFile(configPath, { noEmit: true }, {
      ...ts.sys,
      getCurrentDirectory: () => repository,
      onUnRecoverableConfigFileDiagnostic: diagnostic,
    });
    if (!parsed) return result();
    // A new, virtual root is supplied below, so an otherwise empty project's
    // "no inputs found" diagnostic is not an error. All other config errors are.
    parsed.errors.filter((entry) => entry.code !== 18003).forEach(diagnostic);
    parsed.options.configFile?.parseDiagnostics.forEach(diagnostic);
    // Keep the non-enumerable configFile metadata supplied by TypeScript.
    const options = parsed.options;
    const host = ts.createCompilerHost(options, true);
    const originalReadFile = host.readFile.bind(host);
    const originalFileExists = host.fileExists.bind(host);
    const originalDirectoryExists = host.directoryExists?.bind(host);
    host.getCurrentDirectory = () => repository;
    host.readFile = (name) => isVirtual(name) ? source : originalReadFile(name);
    host.fileExists = (name) => isVirtual(name) || originalFileExists(name);
    // Support imports resolving back to the virtual file in a not-yet-created directory.
    host.directoryExists = (name) => {
      const relative = path.relative(canonical(name), canonical(path.dirname(intendedPath)));
      return relative === '' || (!relative.startsWith(`..${path.sep}`) && relative !== '..'
        && !path.isAbsolute(relative)) || Boolean(originalDirectoryExists?.(name));
    };
    // Never emit, even if a future refactor accidentally requests it.
    host.writeFile = () => { throw new Error('Emitting files is forbidden during validation.'); };
    const program = ts.createProgram({
      rootNames: [...parsed.fileNames.filter((name) => !isVirtual(name)), intendedPath],
      options,
      projectReferences: parsed.projectReferences,
      host,
    });
    const generated = program.getSourceFile(intendedPath);
    program.getOptionsDiagnostics().forEach(diagnostic);
    program.getGlobalDiagnostics().forEach(diagnostic);
    program.getConfigFileParsingDiagnostics().forEach(diagnostic);
    if (!generated) {
      add(issues, 'The intended file could not be loaded as TypeScript source.');
      return result();
    }
    program.getSyntacticDiagnostics(generated).forEach(diagnostic);
    program.getSemanticDiagnostics(generated).forEach(diagnostic);
    inspectDraft(ts, generated, program.getTypeChecker(), (message) => add(issues, message),
      (message) => add(warnings, message));
  } catch (error) {
    add(issues, `Static validation could not complete: ${error instanceof Error ? error.message : String(error)}`);
  }
  return result();
}

function inspectDraft(ts, sourceFile, checker, issue, warn) {
  const unwrap = (node) => {
    while (node && (ts.isParenthesizedExpression(node) || ts.isAsExpression(node)
      || ts.isNonNullExpression(node) || ts.isTypeAssertionExpression(node))) node = node.expression;
    return node;
  };
  const chain = (node) => {
    node = unwrap(node);
    if (ts.isIdentifier(node)) return [node.text];
    if (ts.isPropertyAccessExpression(node)) return [...chain(node.expression), node.name.text];
    if (ts.isElementAccessExpression(node) && node.argumentExpression
        && (ts.isStringLiteral(node.argumentExpression) || ts.isNoSubstitutionTemplateLiteral(node.argumentExpression))) {
      return [...chain(node.expression), node.argumentExpression.text];
    }
    return [];
  };
  const at = (node) => `Line ${sourceFile.getLineAndCharacterOfPosition(node.getStart(sourceFile)).line + 1}`;
  const comments = new Set();
  let executableTests = 0;
  const inspectText = (text, location) => {
    if (/\b(?:TODO|FIXME)\b|\bNEEDS[\s_-]+INPUT\b/i.test(text)) {
      issue(`${location}: unresolved TODO/FIXME or NEEDS INPUT marker; implement the missing behavior/assertion or supply the missing input.`);
    }
    if (/\bGUESS\b/i.test(text)) warn(`${location}: GUESS value requires review against the intended scenario.`);
    const content = text.replace(/^\s*(?:\/[/\*]|\*)\s?/gm, '');
    if (/\b(?:not[\s_-]+implemented|replace[\s_-]+me|your[\s_-]+(?:code|selector)[\s_-]+here)\b|\b(?:implementation|assertion)[\s_-]+goes[\s_-]+here\b|\b(?:missing|unimplemented|unsupported)\s+(?:framework\s+)?(?:functionality|assertions?|behavior)\b|^\s*(?:@automation-agent-data\s*:\s*\w+\s*=\s*)?placeholder\b|<\s*(?:placeholder|insert[^>]*|replace[^>]*)>/im.test(content)) {
      issue(`${location}: unresolved implementation placeholder.`);
    }
  };
  const inspectComment = (range) => {
    if (comments.has(range.pos)) return;
    comments.add(range.pos);
    inspectText(sourceFile.text.slice(range.pos, range.end),
      `Line ${sourceFile.getLineAndCharacterOfPosition(range.pos).line + 1}`);
  };
  const visit = (node) => {
    (ts.getLeadingCommentRanges(sourceFile.text, node.pos) ?? []).forEach(inspectComment);
    (ts.getTrailingCommentRanges(sourceFile.text, node.end) ?? []).forEach(inspectComment);
    if (ts.isStringLiteralLike(node) || ts.isTemplateHead(node) || ts.isTemplateMiddle(node)
        || ts.isTemplateTail(node)) inspectText(node.text, at(node));
    if (ts.isPropertyAccessExpression(node) || ts.isElementAccessExpression(node)) {
      const parts = chain(node);
      if (parts[0] === 'test' && parts.slice(1).some((part) => ['skip', 'fixme', 'only'].includes(part))) {
        issue(`${at(node)}: forbidden Playwright modifier ${parts.join('.')}; tests and suites must not be skipped, disabled, or focused.`);
      }
    }
    if (ts.isCallExpression(node) && chain(node.expression).join('.') === 'test') {
      let callback = unwrap(node.arguments.at(-1));
      if (callback && ts.isIdentifier(callback)) {
        const declaration = checker.getSymbolAtLocation(callback)?.valueDeclaration;
        callback = declaration && ts.isVariableDeclaration(declaration)
          ? unwrap(declaration.initializer) : declaration;
      }
      if (node.arguments.length >= 2 && callback
          && (ts.isArrowFunction(callback) || ts.isFunctionExpression(callback) || ts.isFunctionDeclaration(callback))
          && callback.body && (!ts.isBlock(callback.body) || callback.body.statements.some((statement) => !ts.isEmptyStatement(statement)))) {
        executableTests += 1;
      } else {
        issue(`${at(node)}: test requires an implemented callback, not an empty body or unresolved callback.`);
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(sourceFile);
  if (!executableTests) issue('No executable Playwright test(...) with an implemented callback was found in the AST.');
}
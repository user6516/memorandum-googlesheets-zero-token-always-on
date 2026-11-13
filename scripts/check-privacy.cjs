// 基础发布前检查。只输出相对文件名及问题类型，不打印匹配到的秘密值。
const fs = require('node:fs');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const ignoredDirectories = new Set(['.git', 'node_modules', 'coverage', 'dist']);
const problems = [];
let checked = 0;
function walk(directory) {
  for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
    const file = path.join(directory, entry.name);
    const relative = path.relative(root, file).split(path.sep).join('/');
    if (entry.isSymbolicLink()) { problems.push([relative, '不支持符号链接，请人工检查']); continue; }
    if (entry.isDirectory()) { if (!ignoredDirectories.has(entry.name)) walk(file); continue; }
    checked++;
    if (/^(?:\.env(?:\..+)?|\.clasprc\.json|\.clasp\.json)$/.test(entry.name) && entry.name !== '.env.example') {
      problems.push([relative, '本地凭据或账号绑定文件']);
    }
    if (/\.(?:png|jpe?g|gif|pdf|xlsx?|csv|zip|pem|key)$/i.test(entry.name)) {
      problems.push([relative, '不应直接发布的个人数据或二进制文件']); continue;
    }
    const text = fs.readFileSync(file, 'utf8');
    const emails = text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) || [];
    if (emails.some(value => !/@example\.(?:com|org|net)$/i.test(value))) problems.push([relative, '非示例邮箱']);
    if (/(?<![A-Za-z])[A-Za-z]:[\\/]|\/(?:Users|home)\/[A-Za-z0-9_-]+\//.test(text)) problems.push([relative, '个人绝对路径']);
    if (/https:\/\/(?:docs\.google\.com\/spreadsheets\/d|script\.google\.com\/d)\/[A-Za-z0-9_-]+/.test(text)) problems.push([relative, '私有表格或脚本链接']);
    if (/AIza[0-9A-Za-z_-]{25,}|gh[pousr]_[A-Za-z0-9]{20,}|-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----/.test(text)) problems.push([relative, '疑似访问令牌或私钥']);
    if (/"(?:scriptId|spreadsheetId|client_secret|private_key)"\s*:\s*"[^\"]+"/.test(text)) problems.push([relative, '服务标识或凭据配置']);
    if (relative === 'src/Code.gs' && !/recipient:\s*'YOUR_EMAIL_HERE'/.test(text)) problems.push([relative, '发布版必须保留收件人占位符']);
  }
}
walk(root);
if (problems.length) {
  problems.forEach(([file, reason]) => console.error(`${file}: ${reason}`));
  process.exitCode = 1;
} else console.log(`Privacy check passed: ${checked} files; only placeholders and example email addresses found.`);

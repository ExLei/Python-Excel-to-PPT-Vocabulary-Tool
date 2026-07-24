# 提交信息生成规则

## 格式

```
<type>(<scope>): <中文描述>
```

## 类型（type）

| type | 用途 |
|------|------|
| feat | 新功能 |
| fix | 修复缺陷 |
| refactor | 重构（不改行为） |
| test | 测试 |
| chore | 构建、依赖、配置等杂项 |
| docs | 文档 |

## 范围（scope）

`diag` `pipeline` `export` `reader` `generator` `template` `cli` `gui`

## 描述要求

- 中文，简洁，祈使语气（"添加"而非"添加了"）
- 不超过 50 字
- 不以句号结尾

## 示例

```
feat(pipeline): 添加字体自动检测
fix(export): 修复特殊字符导致空白 PNG
refactor(reader): 简化列映射逻辑
chore: 更新 .gitignore 忽略 .codegraph
docs: 补充诊断系统使用说明
```

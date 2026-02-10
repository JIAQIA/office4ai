# Auto Certificate Management Spec

> 自动证书管理方案 — 让 `uvx` 用户零手动配置即可使用 HTTPS Socket.IO

**状态**: Draft
**日期**: 2026-02-10

---

## 1. 背景与动机

Office Add-In 运行在 WebView (macOS WKWebView / Windows Edge WebView2) 中，通过 Socket.IO 连接本地 Python 服务。WebView 强制要求 HTTPS 连接信任系统级证书存储，因此需要：

1. 生成自签名 CA + 服务证书
2. 将 CA 安装到系统信任存储
3. 服务启动时加载证书

当前方案依赖用户手动执行 `mkcert` 命令，对 `uvx` 分发场景不友好。

## 2. 目标平台

| 平台 | WebView 引擎 | 证书信任存储 | 信任安装命令 | 信任移除命令 |
|------|-------------|-------------|-------------|-------------|
| **macOS** | WKWebView | 系统钥匙串 (Keychain) | `sudo security add-trusted-cert` | `sudo security remove-trusted-cert` |
| **Windows** | Edge WebView2 | Windows 证书存储 | `certutil -addstore Root` (UAC) | `certutil -delstore Root` (UAC) |

两个平台的共同点：WebView **只信任系统级证书存储**中的 CA，不支持应用级信任配置。

## 3. CLI 设计

将当前单一入口 `office4ai-mcp` 改为子命令模式：

```
office4ai-mcp serve          # 启动 MCP Server（默认行为，裸调用等价）
office4ai-mcp setup          # 生成证书 + 安装 CA 到系统信任存储
office4ai-mcp cleanup        # 删除证书 + 从系统信任存储移除 CA
```

### 3.1 `office4ai-mcp`（裸调用）

等价于 `office4ai-mcp serve`，保持向后兼容。

### 3.2 `office4ai-mcp setup`

首次使用前运行一次。交互流程：

```
$ office4ai-mcp setup

🔐 Office4AI Certificate Setup
================================
This will:
  1. Generate a local CA certificate (with Name Constraints: localhost/127.0.0.1 only)
  2. Generate a server certificate for localhost and 127.0.0.1
  3. Install the CA into your system trust store (requires admin privileges)

Certificate location: ~/.office4ai/certs/

Proceed? [y/N]: y

✅ CA certificate generated (valid 10 years)
✅ Server certificate generated (valid 825 days)
🔑 Installing CA to system trust store...
   [macOS] sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain ca.pem
   [Windows] certutil -addstore Root ca.pem
✅ CA installed to system trust store
✅ Setup complete! You can now start the server with: office4ai-mcp serve
```

**行为规则**：
- 如果 `~/.office4ai/certs/` 已存在有效证书（未过期） → **跳过**，提示已就绪
- 如果证书已过期或不存在 → 生成新证书并安装 CA
- 不提供 `--yes` flag（安装 CA 到系统是安全敏感操作，强制交互确认）

### 3.3 `office4ai-mcp cleanup`

```
$ office4ai-mcp cleanup

🧹 Office4AI Certificate Cleanup
===================================
This will:
  1. Remove CA from system trust store (requires admin privileges)
  2. Delete all certificate files from ~/.office4ai/certs/

Proceed? [y/N]: y

✅ CA removed from system trust store
✅ Certificate files deleted
✅ Cleanup complete
```

### 3.4 `office4ai-mcp serve`

启动时检查证书：

```
# 无证书 → Hard Fail
$ office4ai-mcp serve
❌ SSL certificates not found at ~/.office4ai/certs/
   Run `office4ai-mcp setup` first to generate and install certificates.
   (exit code 1)

# 证书已过期 → Hard Fail
$ office4ai-mcp serve
❌ Server certificate expired (2026-01-15)
   Run `office4ai-mcp setup` to regenerate certificates.
   (exit code 1)

# 正常启动
$ office4ai-mcp serve
✅ Loaded certificates from ~/.office4ai/certs/
   CA: valid until 2036-02-10
   Server cert: valid until 2028-05-06
...
```

**不支持 HTTP-only fallback**。Add-In 必须走 HTTPS，HTTP 模式无实际用途。

## 4. 证书生成方案

### 4.1 技术选型：Pure Python

使用 `cryptography` 库（已是常见间接依赖），零外部依赖。

**不依赖 mkcert**：避免要求用户安装非 Python 二进制。

### 4.2 证书体系

```
~/.office4ai/certs/
├── ca.pem              # CA 证书（安装到系统信任存储）
├── ca-key.pem          # CA 私钥（权限 0600）
├── cert.pem            # 服务器证书
└── key.pem             # 服务器私钥（权限 0600）
```

### 4.3 CA 证书

| 属性 | 值 |
|------|----|
| Subject | `CN=Office4AI Local CA` |
| 有效期 | **10 年** |
| Key Usage | Certificate Sign, CRL Sign |
| Basic Constraints | CA:TRUE, pathlen:0 |
| **Name Constraints** | permitted: DNS:localhost, IP:127.0.0.1, IP:::1 |
| Key Size | RSA 2048 (或 EC P-256) |

**Name Constraints 扩展**：即使 CA 私钥泄露，攻击者也无法签发 `google.com` 等外部域名的证书。CA 仅被允许为 localhost/127.0.0.1/::1 签发证书。这是 mkcert 所不具备的安全增强。

CA 有效期 10 年在两个平台上均无限制问题：
- macOS 钥匙串对本地安装的 CA 证书本身无有效期上限
- Windows 证书存储同样不限制本地 CA 的有效期

### 4.4 服务器证书

| 属性 | 值 | 说明 |
|------|----|------|
| Subject | `CN=localhost` | |
| 有效期 | **825 天** | macOS WKWebView 对本地 CA 签发的证书限制上限 |
| SAN | DNS:localhost, IP:127.0.0.1, IP:::1 | |
| Key Usage | Digital Signature, Key Encipherment | |
| Extended Key Usage | Server Authentication | |

#### 有效期跨平台分析

| 平台 | 本地 CA 签发的服务证书有效期限制 | 来源 |
|------|-------------------------------|------|
| **macOS** (WKWebView) | **825 天**（≈2年3个月） | Apple 对 user-added CA 签发证书的强制限制，超过此值 Safari/WKWebView 会拒绝信任 |
| **Windows** (Edge WebView2) | **无严格上限** | Edge WebView2 走 Windows 证书存储，不对本地 CA 签发的证书施加额外有效期限制 |

**统一取 825 天**：取两平台约束的交集（即 macOS 的更严格限制），确保同一套证书在两个平台上均可正常工作。Windows 上 825 天同样有效，无兼容问题。

> 注：Apple 对公共 CA 的 398 天限制（及未来的 200→100→47 天缩短计划）**不适用于**本地安装的 CA。本地 CA 的 825 天限制是独立的、目前稳定的策略。

### 4.5 私钥安全

| 措施 | macOS | Windows |
|------|-------|---------|
| 私钥文件权限 | `chmod 0600`（仅所有者可读写） | `icacls` 设置仅当前用户 Full Control，移除 Inherited/Everyone |
| CA 私钥保留 | 保留在磁盘，用于未来证书续期 | 同左 |
| 证书目录权限 | `chmod 0700 ~/.office4ai/certs/` | `icacls` 限制目录访问 |

## 5. 系统信任安装（跨平台）

### 5.1 macOS

**安装 CA 到系统钥匙串**：
```bash
sudo security add-trusted-cert -d -r trustRoot \
  -k /Library/Keychains/System.keychain \
  ~/.office4ai/certs/ca.pem
```

**移除 CA**：
```bash
sudo security remove-trusted-cert -d ~/.office4ai/certs/ca.pem
```

**检查 CA 是否已安装**：
```bash
security find-certificate -c "Office4AI Local CA" /Library/Keychains/System.keychain
```

| 要点 | 说明 |
|------|------|
| 权限要求 | `sudo`（终端弹出密码输入框） |
| 信任范围 | 全系统（所有用户、所有应用、WKWebView） |
| stdin 处理 | **不捕获 stdin**，让终端直接传递密码输入给 `sudo` |

### 5.2 Windows

**安装 CA 到受信任的根证书颁发机构**：
```cmd
certutil -addstore Root "%USERPROFILE%\.office4ai\certs\ca.pem"
```

**移除 CA**：
```cmd
certutil -delstore Root "Office4AI Local CA"
```

**检查 CA 是否已安装**：
```cmd
certutil -store Root "Office4AI Local CA"
```

| 要点 | 说明 |
|------|------|
| 权限要求 | UAC 弹窗（用户点击"是"） |
| 信任范围 | 当前用户或全机器（取决于 UAC 授权），Edge WebView2 可读取 |
| stdin 处理 | UAC 由系统接管，Python 进程无需特殊处理 |

### 5.3 跨平台实现策略

```python
# trust_store.py 伪代码结构

class TrustStore(ABC):
    @abstractmethod
    def install(self, ca_cert_path: Path) -> bool: ...
    @abstractmethod
    def uninstall(self, ca_cn: str, ca_cert_path: Path) -> bool: ...
    @abstractmethod
    def is_installed(self, ca_cn: str) -> bool: ...

class MacOSTrustStore(TrustStore): ...
class WindowsTrustStore(TrustStore): ...

def get_trust_store() -> TrustStore:
    """根据 platform.system() 返回对应实现"""
```

**共同实现注意**：
- 通过 `subprocess.run()` 调用平台命令
- 捕获返回码和 stderr，失败时给出**平台特定的**错误信息和手动操作指引
- 安装前先调用 `is_installed()` 检查，避免重复安装
- 失败时的 fallback 消息包含手动执行的完整命令，让用户可以 copy-paste

## 6. 证书路径发现（跨平台）

### 6.1 默认路径

| 平台 | 默认路径 |
|------|---------|
| macOS | `~/.office4ai/certs/` |
| Windows | `%USERPROFILE%\.office4ai\certs\` |

两个平台统一使用用户 home 目录下的 `.office4ai/certs/`，通过 `Path.home() / ".office4ai" / "certs"` 跨平台获取。

### 6.2 查找优先级

1. 环境变量 `OFFICE4AI_CERT_DIR` → 使用指定目录
2. 默认路径 `~/.office4ai/certs/`

这允许：
- 开发者使用已有证书（如 mkcert 生成的）放在自定义位置
- CI/Docker 环境通过环境变量注入证书路径

`OfficeWorkspace` 不再硬编码 `certs/` 相对路径，改为接收 `cert_dir` 参数。

## 7. 证书生命周期（跨平台统一）

### 7.1 生命周期状态机

```
                    ┌──────────┐
        setup       │  No Cert │  ← 初始状态 / cleanup 后
                    └────┬─────┘
                         │ office4ai-mcp setup
                         ▼
                    ┌──────────┐
                    │  Valid   │  ← CA (10y) + Server Cert (825d) 均有效
                    └────┬─────┘
                         │ 时间流逝...
                         ▼
              ┌─────────────────────┐
              │ Server Cert Expired │  ← CA 仍有效（10y 远未到期）
              └────────┬────────────┘
                       │ office4ai-mcp setup（自动检测）
                       ▼
              ┌─────────────────────┐
              │ Renew Server Cert   │  复用 CA 私钥重新签发 825d 服务证书
              │ (CA 不变，无需重装    │  无需再次安装 CA 到系统信任存储
              │  系统信任)           │
              └────────┬────────────┘
                       │
                       ▼
                    ┌──────────┐
                    │  Valid   │
                    └──────────┘
```

### 7.2 事件-行为矩阵

| 事件 | 行为 | macOS 特有 | Windows 特有 |
|------|------|-----------|-------------|
| **首次 `setup`** | 生成 CA (10y) + 服务证书 (825d) + 安装 CA | `sudo security add-trusted-cert` | `certutil -addstore Root` (UAC) |
| **`serve` 启动** | 检查证书存在性 + 有效期 → hard fail 或正常启动 | — | — |
| **服务证书过期** | `serve` 报错，提示 `setup` | — | — |
| **重新 `setup`（证书有效）** | 跳过，提示已就绪 | — | — |
| **重新 `setup`（服务证书过期，CA 有效）** | 复用 CA 重新签发服务证书，**无需重装系统信任** | — | — |
| **重新 `setup`（CA 也过期）** | 重新生成 CA + 服务证书 + 重装系统信任 | `sudo` 密码 | UAC 弹窗 |
| **`cleanup`** | 从系统信任存储移除 CA + 删除所有证书文件 | `sudo security remove-trusted-cert` | `certutil -delstore Root` (UAC) |

### 7.3 续期优化

服务证书续期（CA 仍有效时）是最常见的续期场景。此时：
- **只需重新生成** `cert.pem` + `key.pem`
- **不需要重新安装** CA 到系统信任存储（CA 没变，系统已信任）
- **不需要管理员权限**（只是写文件到 `~/.office4ai/certs/`）

这是关键的 UX 优化：10 年内的续期操作对用户完全透明，无需任何特权操作。

## 8. 代码变更范围

### 8.1 新增模块

```
office4ai/
└── certs/                          # 新增：证书管理模块
    ├── __init__.py
    ├── generator.py                # CA + 服务证书生成（cryptography）
    ├── trust_store.py              # 平台信任存储安装/移除
    ├── paths.py                    # 证书路径发现逻辑
    └── validator.py                # 证书有效性检查
```

### 8.2 修改模块

| 文件 | 变更 |
|------|------|
| `office/mcp/server.py` | `main()` 改为子命令路由 (serve/setup/cleanup) |
| `environment/workspace/office_workspace.py` | `start()` 中证书路径从硬编码改为 `paths.get_cert_dir()` |
| `pyproject.toml` | 添加 `cryptography` 依赖 |
| `a2c_smcp/config.py` | 可选：添加 `cert_dir` 配置项 |

### 8.3 不变

- Socket.IO 端口计算逻辑 (`port + 1443`) 不变
- `E2ETestRunner` 保持现有行为（开发环境仍可用 `OFFICE4AI_CERT_DIR` 指向项目 `certs/`）
- MCP transport 逻辑不变

## 9. 依赖变更

```toml
[project]
dependencies = [
    # 已有依赖 ...
    "cryptography>=42.0",   # 新增：证书生成
]
```

`cryptography` 是 Python 生态中最广泛使用的加密库，大部分环境已有间接安装。

## 10. 安全考量

| 风险 | 缓解措施 | 平台适用 |
|------|---------|---------|
| CA 私钥泄露签发任意证书 | **Name Constraints** 限制 CA 只能签 localhost/127.0.0.1/::1 | 两平台 |
| 私钥文件权限过宽 | macOS: `chmod 0600`；Windows: `icacls` 限制为当前用户 | 各自处理 |
| 用户不理解 CA 安装的含义 | setup 命令明确解释操作内容，要求 y/N 确认 | 两平台 |
| 卸载后 CA 残留 | 提供 `cleanup` 命令，分别调用平台移除命令 | 各自处理 |
| Windows UAC 静默拒绝 | 检测 `certutil` 返回码，失败时打印手动命令 | Windows |
| macOS sudo 密码输入失败 | 检测返回码，提示重试或给出手动命令 | macOS |

## 11. 参考来源

- [Apple: Requirements for trusted certificates in iOS 13 and macOS 10.15](https://support.apple.com/en-us/103769) — 本地 CA 签发证书 825 天限制
- [Apple: About upcoming limits on trusted certificates](https://support.apple.com/en-us/102028) — 公共 CA 398 天限制（不适用于本地 CA）
- [Microsoft: Edge browser TLS certificate verification](https://learn.microsoft.com/en-us/deployedge/microsoft-edge-security-cert-verification) — Edge/WebView2 信任本地安装的 CA
- [Michal Špaček: Validity period of HTTPS certificates issued from a user-added CA](https://www.michalspacek.com/validity-period-of-https-certificates-issued-from-a-user-added-ca-is-essentially-2-years) — 825 天限制的详细分析
- [mkcert](https://github.com/FiloSottile/mkcert) — 设计参考（825 天 leaf + 10 年 CA）

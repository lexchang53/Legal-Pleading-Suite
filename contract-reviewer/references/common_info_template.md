# 公司資訊模板

使用契約審核 Skill 時，可以提供公司常用資訊用於自動填充契約。

## 資訊格式

可以透過以下方式提供公司資訊：

### 方式一：提供 .docx 文件

創建一個 Word 文件，包含以下資訊（格式不限，只要包含關鍵資訊即可）：

```
公司名稱：[您的公司名稱]
代表人：[姓名]
登記地址：[詳細地址]
辦公地址：[詳細地址]（如與登記地址相同可省略）

聯絡人資訊：
  姓名：[聯絡人姓名]
  電話：[聯絡電話]
  信箱：[電子信箱]

銀行帳戶資訊：
  開戶行：[銀行名稱]
  帳號：[銀行帳號]
  戶名：[帳戶名稱]
```

### 方式二：直接告訴 Claude

在對話中直接提供資訊，例如：

```
公司資訊：
- 公司名稱：XX科技有限公司
- 代表人：張三
- 地址：台北市XX區XX路XX號
- 聯絡人：李四，電話 138XXXXXXXX
- 開戶行：XX銀行XX支行，帳號：XXXXXXXXX
```

### 方式三：提供 JSON 文件

創建 `company_info.json` 文件：

```json
{
  "company_name": "XX科技有限公司",
  "legal_representative": "張三",
  "registered_address": "台北市XX區XX路XX號",
  "office_address": "台北市XX區XX路XX號",
  "contact_person": "李四",
  "contact_phone": "138XXXXXXXX",
  "contact_email": "contact@example.com",
  "bank_name": "XX銀行XX支行",
  "bank_account": "XXXXXXXXXXXXXXXXX",
  "account_name": "XX科技有限公司"
}
```

## 支持填充的欄位

| 欄位類型 | 欄位名稱 | 說明 |
|---------|---------|------|
| 公司基本資訊 | company_name | 公司全稱 |
| | legal_representative | 代表人姓名 |
| | registered_address | 登記地址 |
| | office_address | 辦公地址 |
| 聯絡人資訊 | contact_person | 聯絡人姓名 |
| | contact_phone | 聯絡電話 |
| | contact_email | 電子信箱 |
| 銀行資訊 | bank_name | 開戶銀行 |
| | bank_account | 銀行帳號 |
| | account_name | 帳戶名稱 |

## 契約中的占位符識別

系統會自動識別契約中的以下占位符並嘗試填充：

- `[乙方公司名稱]`、`[公司名稱]` → 公司名稱
- `[代表人]` → 代表人
- `[地址]`、`[登記地址]` → 地址
- `[聯絡人]` → 聯絡人
- `[電話]`、`[聯絡電話]` → 電話
- `[信箱]` → 信箱
- `[開戶行]` → 開戶行
- `[帳號]` → 銀行帳號
- 以 `：` 或 `:` 結尾的空白行也會嘗試匹配填充

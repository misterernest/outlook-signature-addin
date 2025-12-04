# Outlook Signature Add-in (WWG)

This project contains a custom Outlook Web Add-in that allows companies to automate the
insertion of HTML email signatures for users in Microsoft 365.

The add-in loads a hosted HTML signature template and inserts it directly into the body of
a new email with a single click. Ideal for organizations that want centralized control
without requiring client-side configuration.

---

## ✨ Features

- One-click signature insertion in Outlook Web, Outlook New, and Outlook Classic
- Fully HTML-based (supports images, styling, external assets)
- Hosted on a secure HTTPS subdomain  
  Example: `https://services.willworldglobal.com/outlook/signature/`
- Custom signature templates (supports placeholders and dynamic fields)
- Lightweight and fast (no backend required)
- Ready for enterprise deployment using Microsoft 365 *Integrated Apps*

---

## 📁 Project Structure

```
src/
├─ manifest.xml
└─ outlook/
	└─ signature/
		├─ index.html
		├─ script.js
		└─ signature-test.html
tools/
└─ generate-signatures.ps1 # Helper tool for generating HTML signatures
docs/
└─ setup-outlook.md # Instructions for installing the add-in
```

---

## 🚀 Installation (for administrators)

1. Host the HTML/JS files on a secure HTTPS site  
	Example folder:  
	`https://services.willworldglobal.com/outlook/signature/`

2. Update `manifest.xml` and set:
	```xml
	<SourceLocation DefaultValue="https://services.yourdomain.com/outlook/signature/index.html" />
	```

3. Go to Microsoft 365 Admin Center → Integrated Apps → Upload custom apps

4. Upload the `manifest.xml` and deploy to either:
	- Only yourself (for testing), or
	- The entire organization

5. Open Outlook Web → New message → Open the add-in from the right-side panel.

🧪 Development Mode

For local development, you can use:

- Live Server
- WAMP/Apache
- VS Code Web Server
- Any local environment

Note: Outlook requires HTTPS when deploying through Microsoft 365 admin center.

🛡️ License

MIT License — You may freely use, modify, and adapt this project commercially or privately.

👤 Author

Ernesto (misterernest)
Will World Global

---

# ✅ **2. .gitignore recomendado**

Selecciona este en GitHub:

**Template: “Node”**

Luego agrega estas líneas manualmente:

```
out/
*.log
.vscode/
*.ps1xml
Thumbs.db
.DS_Store
```

---

# ✅ **3. LICENSE (MIT)**

Selecciona **“Add license → MIT License”** en GitHub  
O si lo quieres pegar manualmente:

```markdown
MIT License

Copyright (c) 2025 Ernesto

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software.
```

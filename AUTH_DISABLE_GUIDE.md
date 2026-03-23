# How to Enable / Disable Login

## Current state: Login is **DISABLED**

The main sheet page opens directly without requiring login.

---

## To Disable Login (no password required)

1. Open `config.json`
2. Set `"disabled": true` inside the `auth` object:

```json
{
  "auth": {
    "disabled": true,
    "users": { ... }
  }
}
```

3. Restart the server (`npm start` or `node server.js`)

---

## To Enable Login (password required)

1. Open `config.json`
2. Set `"disabled": false` or remove the `"disabled"` line:

```json
{
  "auth": {
    "disabled": false,
    "users": {
      "gourav.singh@thecodeinsight.com": "CodeInsight@123",
      "jitesh.kumar@thecodeinsight.com": "admin@123"
    }
  }
}
```

3. Restart the server

---

## Quick reference

| `auth.disabled` | Behavior |
|----------------|----------|
| `true`         | No login — main sheet opens at http://localhost:PORT/ |
| `false` or omitted | Login required — redirected to /login if not authenticated |

# Setting up Gmail access for gmail_extractor.py

This has to be done once, by you, in a browser — I can't create a Google
Cloud project or OAuth credentials on your behalf. It takes about 5 minutes.

Use your **own personal Gmail account** for this test (not the library's —
you don't have access to that yet). The goal right now is just to prove the
extraction code works; pointing it at the real library inbox is a later
step once the library grants access.

## 1. Create a Google Cloud project
1. Go to https://console.cloud.google.com/
2. Click the project dropdown (top left) → **New Project**
3. Name it anything (e.g. "library-ai-checker-test") → **Create**

## 2. Enable the Gmail API
1. With your new project selected, go to
   https://console.cloud.google.com/apis/library/gmail.googleapis.com
2. Click **Enable**

## 3. Configure the OAuth consent screen
1. Go to https://console.cloud.google.com/apis/credentials/consent
2. Choose **External** → **Create**
3. Fill in required fields (app name, your email for support + developer
   contact) → **Save and Continue** through the remaining steps
4. On the **Test users** step, add your own Gmail address → **Save**
   (while the app is in "Testing" mode, only test users you list can
   authorize it — that's fine for prototyping)

## 4. Create OAuth credentials
1. Go to https://console.cloud.google.com/apis/credentials
2. Click **Create Credentials** → **OAuth client ID**
3. Application type: **Desktop app** → name it anything → **Create**
4. Click **Download JSON** on the credential you just created
5. Rename the downloaded file to `credentials.json` and put it in this
   project folder (same folder as `gmail_extractor.py`)

## 5. Install the new dependencies
```
pip install -r requirements.txt
```

## 6. Authorize and test
```
python gmail_auth.py
```
A browser window opens asking you to sign in and approve access. Approve
it — you'll see a warning that the app is "unverified," which is expected
for a project still in testing; click **Advanced** → **Go to (your app
name)** to continue. Once approved, it saves a `token.json` locally so you
won't need to repeat this every run.

You should see:
```
Authorized as: your.email@gmail.com
```

## 7. Fetch some real emails
```
python gmail_extractor.py --max 10
```
This pulls your 10 most recent emails (whatever's in your inbox — to test
realistically, email yourself a couple of test messages with a PDF
attached first). Check:
- `email_input/` for saved attachments
- `output/email_metadata.xlsx` for the subject/from/to/cc/body log

## Notes
- `credentials.json` and `token.json` both grant access to your mailbox —
  add them to `.gitignore` and never commit them.
- The scope used is **read-only** (`gmail.readonly`) — this script cannot
  send, delete, or modify any mail, only read it.
- When you're ready to point this at the real library inbox, the library
  account owner (or your IT admin) needs to either add that account as a
  test user on this OAuth app, or set up a proper Workspace service
  account with domain-wide delegation — a slightly different flow we can
  set up once that access is granted.

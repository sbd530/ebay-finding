# :wrench: Installation

- Install Node.js LTS version.
  - https://nodejs.org/ko/download/
  - Check version by `node -v` and `npm -v`

- Git clone or download code.

```bash
$ npm install
```

# :computer: Execution

- Create `.env` file at the root directory to execute app.
  - Refer to `.env.sample` file.

```
.
├── .env.sample
├── .env
...
```

- Run the app.

```bash
$ npm run start
```

# :books: To get data

- Check out data files in `data` diretory.
  - Files are formatted in CSV format.

- If result data doesn't exist for your keywords, the app may print out `Result is empty`.

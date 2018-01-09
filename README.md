

### made add-in in office
```
npm install 

mv ~/Library/Containers/com.microsoft.Word/Data/Documents ~/Library/Containers/com.microsoft.Word/Data/documents

mkdir ~/Library/Containers/com.microsoft.Word/Data/documents/wef

cp my-office-add-in-manifest.xml ~/Library/Containers/com.microsoft.Word/Data/documents/wef

```
### Issues on Certificate

The project is constucted by [the official instruction](https://docs.microsoft.com/en-us/office/dev/add-ins/quickstarts/word-quickstart?tabs=visual-studio-code)

We don't upload certs/ directory, which contains some confidential data.
Therefore, you may need create your own certificate for localhost server. If you encounter Subject Alternative Name problem, try [this one](https://github.com/OfficeDev/generator-office/issues/274).

This repo has already fixed `Cannot GET /` by adding `--server` in `browser-sync start --config bsconfig.json` .


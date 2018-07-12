# configuration 

to list reports
```javascript
{

  "action": "list_reports",
  "username": "foo@bar.com",
  "#password": "secret"
}
```

to download reports
```javascript
{

  "username": "foo@bar.com",
  "#password": "secret"
  "action": "download_reports",
  "reports": [
    {
      "report_id": 12345,
      "filename":  "friendly_name_less_than_64_chars"
    }
  ]
}
```

Excels are converted to csv with `\u0001` [delimiter](https://www.fileformat.info/info/unicode/char/0001/index.htm)  

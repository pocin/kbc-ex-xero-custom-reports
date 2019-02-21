# configuration 

to list reports
```javascript
{

  "action": "list_reports",
  "username": "foo@bar.com",
  "account_id": "!abc12",
  "#password": "secret"
}
```

to download reports
```javascript
{

  "username": "foo@bar.com",
  "#password": "secret"
  "account_id" "!abc12",
  "action": "download_reports",
  "reports": [
    {
      "report_id": 12345,
      "filename":  "friendly_name_less_than_64_chars",
      "from_date": "last month",
      "to_date": "today",
      "delay_seconds": 15,
    }
  ]
}
```

to use direct url for base reports (not custom!)
```javascript
{
	"username": "foo@bar.com",
	"#password": "secret",
	"account_id": "!abc12"
	"action": "direct_url",
	"direct_url": "https://go.xero.com/Reports/Report.aspx?reportId=abcdef-[truncated]&statement=123456-[truncated]&reportClass=ProfitAndLoss"
}
```

the optional `delay_seconds` parameter controls how long to wait after updating the date range. If the report large and you get errors like 
```
selenium.common.exceptions.InvalidElementStateException: Message: invalid element state
(Session info: headless chrome=70.0.3538.77)
(Driver info: chromedriver=2.40.565383 (76257d1ab79276b2d53ee976b2c3e3b9f335cde7),platform=Linux 4.9.32-15.41.amzn1.x86_64 x86_64)
```

try increasing the value, 15 seconds should be enough for most cases (the default), but feel free to experiment.


## Date ranges
Choose 
```
    "from_date": "timestamp",
    "to_date": "timestamp",
```

if the date range in the ui looks like this

![date_range_between](./media/date_range_between.png)

or

```
    "to_date": "timestamp" 
```
if the date range in the ui looks like this
![date_range_until](./media/date_range_until.png)


if you omit the date range, the default value (defined in the UI) will be used

The value can be either a human readable string or a string in format `23 Jun 2018`

The human readable strings can be 
- `beginning of previous month`
- `beginning of previous month - 2 months` (to look 3 months back)
- `end of previous month` (to look 3 months back)

Always check the logs to make sure the date is converted as expected

## Account id
Get account_id by logging in to the UI, clicking reports and copying it from the url

![How to get account_id from xero ui](./media/xero_account_id_ui.png)

Excels are converted to csv and uploaded to keboola with with [\u0001](https://www.fileformat.info/info/unicode/char/0001/index.htm) delimiter so that all columns are squashed into one called "data"
# Development
```
$ mkdir data
$ # make a file #data/config.json
$ docker-compose build dev
$ docker-compose run --rm dev bash
$ python3 main.py
$ #check  /tmp folder in container for screenshots
```

# License
This project is licensed under [GLWTPL](https://github.com/me-shaon/GLWTPL/blob/master/LICENSE) License

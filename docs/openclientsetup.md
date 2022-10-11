# OpenClient Setup
    
## Prerequisites

- Node.js 12 ( Node.js 10 would also work, but 12 is used when develop & test)
- ngrok (Used to start HTTPS proxy server for autodiscover)
- Postman for verification

## Development 
    
### git

### Test nsfs
- emuppet.nsf
- bmuppet.nsf
- DemoMail.nsf

### Test Users
- Ernie Muppet
- Bert Muppet
- John Doe


### quattro.rocks environment
		id_quattro ssh id file

        Copy ews code to quattro
		`rsync -r . quattro:/opt/hcl/ews`

        ssh to quattro environment
		`ssh quattro`

		Process management
        pm2
			`pm2 list`
			`pm2 stop ews`
			`pm2 start ews`

		restart quattro server and keep  (it will disconnect ssh and take a couple of minutes)
		`sudo shutdown -r now`

        Users

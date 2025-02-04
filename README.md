# Geeral info

* You can use Selenium + [WinAppDriver](https://github.com/microsoft/WinAppDriver) to automate application interactions.
* Have to enable developer mode in Windows
* Have to use python 3.7
* To make it work you have to first start the WinAppDriver.exe application and then run your automation script.
* Default port is 4727
* Windows Application Driver install path - `C:\Program Files (x86)\Windows Application Driver`
* For windows GUI application [walk](https://github.com/lxn/walk) is used.


# How to run

0. Prepare your environment:
    * Download and install [WinAppDriver](https://github.com/microsoft/WinAppDriver)
    * [Enable developer mode in Windows](https://learn.microsoft.com/en-us/windows/apps/get-started/enable-your-device-for-development)
    * Make sure you have python 3.7 installed (including the `py` launcher). To
      check if you have it installed, open cmd.exe and type `py -3.7` (you can
      quit the python interpreter by typing `quit()` and hitting enter).
    * Install python requirements by openining cmd.exe in the project directory
      and running `py -3.7 -m pip install -r requirements.txt`
    * Create a folder in Outlook and a rule that moves all incomming meeting
      requests to that folder.
    * For that new folder make sure to [disable date headers](https://answers.microsoft.com/en-us/outlook_com/forum/all/how-do-i-remove-the-date-grouping-in-the-new/e3267590-6abd-4545-b8c4-ddf9317dbbd7)
1. Fill in the information in the config file `config.ini`.

    NOTE: Make sure to set `RunWinAppDriver` to `yes` if you want the script to
    start it automatically for you
2. To run the application execute the following command: `py -3.7 outlook.py`

    IMPORTANT: Outlook must be running before executing the script. Also all
    other opened outlook windows must be closed (notification window, email
    composer window etc.)

    NOTE: While the script is running do not interact with the PC


The script will go through each meeting request in the specified folder and
forward them to the configured email address. The script will store any forwarded
meeting so next time the script runs it does not forward it again. In case the
meeting forwarding is disabled, the meeting wont be forwarded and the script will
print to the console that the meeting forwarding was not successfull.

# Build the GUI

All these steps must be done on windows

1. Make sure to install [rsrc](https://github.com/akavel/rsrc)
```
go install github.com/akavel/rsrc@latest
```

2. Run the following commands

```
rsrc -manifest test.manifest -o rsrc.syso
go build -ldflags="-H windowsgui"
```

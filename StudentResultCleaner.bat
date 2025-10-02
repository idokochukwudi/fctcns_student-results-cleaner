@echo off
     echo Setting up StudentResultCleaner...

     :: Ensure WSL is running
     wsl -e bash -c "echo WSL is running"
     if %ERRORLEVEL% NEQ 0 (
         echo Error: WSL failed to start. Run 'wsl --shutdown' and 'wsl' in PowerShell as Administrator.
         pause
         exit /b
     )

     :: Get WSL IP
     for /f "tokens=2 delims= " %%i in ('wsl ip addr show eth0 ^| findstr inet ^| findstr /v inet6') do set WSL_IP=%%i
     set WSL_IP=%WSL_IP:~0,-3%
     if "%WSL_IP%"=="" (
         echo Error: Could not detect WSL IP. Check WSL network configuration.
         pause
         exit /b
     )
     echo Detected WSL IP: %WSL_IP%

     :: Update port forwarding
     netsh interface portproxy delete v4tov4 listenport=80 listenaddress=0.0.0.0
     netsh interface portproxy delete v4tov4 listenport=5000 listenaddress=0.0.0.0
     netsh interface portproxy add v4tov4 listenport=80 listenaddress=0.0.0.0 connectport=80 connectaddress=%WSL_IP%
     netsh interface portproxy add v4tov4 listenport=5000 listenaddress=0.0.0.0 connectport=5000 connectaddress=%WSL_IP%
     netsh interface portproxy show all

     :: Ensure firewall rules exist
     netsh advfirewall firewall show rule name="StudentResultCleaner" >nul
     if %ERRORLEVEL% NEQ 0 (
         echo Adding firewall rules...
         netsh advfirewall firewall add rule name="StudentResultCleaner" dir=in action=allow protocol=TCP localport=80
         netsh advfirewall firewall add rule name="StudentResultCleaner" dir=in action=allow protocol=TCP localport=5000
     )

     :: Start Nginx
     wsl -e bash -c "sudo service nginx restart"
     wsl -e bash -c "sudo service nginx status | grep running"
     if %ERRORLEVEL% NEQ 0 (
         echo Error: Nginx failed to start. Check /var/log/nginx/error.log in WSL.
         pause
         exit /b
     )

     :: Start Gunicorn via script
     wsl -e bash -c "~/student_result_cleaner/start_gunicorn.sh"
     wsl -e bash -c "sleep 5 && curl http://127.0.0.1:5000 > /dev/null 2>&1"
     if %ERRORLEVEL% NEQ 0 (
         echo Error: Gunicorn failed to start. Check ~/student_result_cleaner/launcher/app.log in WSL.
         pause
         exit /b
     )

     :: Open browser
     start http://192.168.162.155

     echo StudentResultCleaner started. Check browser at http://192.168.162.155
     pause
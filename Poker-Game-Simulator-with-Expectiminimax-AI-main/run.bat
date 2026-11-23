
@echo off

REM ============================================================
REM
REM              POKER AI SIMULATOR
REM
REM              WINDOWS SETUP SCRIPT
REM
REM ============================================================
REM
REM Project:
REM Poker Game Simulator with Expectiminimax
REM
REM Course:
REM CSE440
REM
REM Purpose:
REM This batch file automatically prepares the Windows
REM environment required to execute the Poker AI Simulator.
REM
REM Main Operations:
REM 1. Detect Windows command environment.
REM 2. Locate the installed Python interpreter.
REM 3. Validate Python version.
REM 4. Navigate to the project directory.
REM 5. Create a virtual environment.
REM 6. Activate the virtual environment.
REM 7. Upgrade pip.
REM 8. Install project dependencies.
REM 9. Verify required Python packages.
REM 10. Verify Streamlit availability.
REM 11. Launch the Poker AI Simulator.
REM 12. Handle installation and execution errors.
REM
REM ============================================================


REM ============================================================
REM SECTION 1: INITIAL CONFIGURATION
REM ============================================================

REM Enable local environment variables.

SETLOCAL ENABLEDELAYEDEXPANSION


REM Define the application name.

SET APP_NAME=Poker AI Simulator


REM Define the project title.

SET PROJECT_TITLE=Poker Game Simulator with Expectiminimax


REM Define the minimum supported Python version.

SET REQUIRED_PYTHON_MAJOR=3

SET REQUIRED_PYTHON_MINOR=10


REM Define Python executable.

SET PYTHON_BIN=python


REM Define virtual environment directory.

SET VENV_DIR=.venv


REM Define virtual environment activation script.

SET VENV_ACTIVATE=.venv\Scripts\activate.bat


REM Define requirements file.

SET REQUIREMENTS_FILE=requirements.txt


REM Define Streamlit entry point.

SET MAIN_FILE=main.py


REM Define retry limits.

SET INSTALL_RETRY_LIMIT=2


REM Initialize error tracking variables.

SET SETUP_ERROR=0

SET PYTHON_FOUND=0

SET VENV_CREATED=0

SET DEPENDENCIES_INSTALLED=0

SET STREAMLIT_READY=0


REM ============================================================
REM SECTION 2: DISPLAY APPLICATION INFORMATION
REM ============================================================

CLS

echo.
echo ============================================================
echo.
echo              POKER AI SIMULATOR SETUP
echo.
echo ============================================================
echo.
echo Project: %PROJECT_TITLE%
echo Application: %APP_NAME%
echo Platform: Windows
echo Required Python: 3.10 or newer
echo Environment: %VENV_DIR%
echo Interface: Streamlit
echo.
echo ============================================================
echo.

echo Initializing the setup process...

echo.

REM Display a short setup introduction.

echo This script will automatically configure the required
echo Python environment and install the project dependencies.

echo.

echo Please wait while the environment is being prepared.

echo.


REM ============================================================
REM SECTION 3: CHECK WINDOWS COMMAND ENVIRONMENT
REM ============================================================

echo ============================================================
echo STEP 1: CHECKING WINDOWS ENVIRONMENT
echo ============================================================

REM Verify that the script is running under Windows.

IF NOT DEFINED WINDIR (

    echo Error: Windows environment could not be identified.

    SET SETUP_ERROR=1

    goto :ERROR_HANDLER

)


REM Display the Windows system directory.

echo Windows directory detected:

echo %WINDIR%

echo.


REM Display the current command processor.

echo Command processor:

echo %COMSPEC%

echo.


REM Verify that the command processor is available.

IF NOT EXIST "%COMSPEC%" (

    echo Error: Windows command processor was not found.

    SET SETUP_ERROR=1

    goto :ERROR_HANDLER

)


echo Windows environment check completed successfully.

echo.


REM ============================================================
REM SECTION 4: DETECT PYTHON INSTALLATION
REM ============================================================

echo ============================================================
echo STEP 2: DETECTING PYTHON INSTALLATION
echo ============================================================

REM Check whether Python is available through PATH.

where python >nul 2>nul

IF ERRORLEVEL 1 (

    echo Python executable was not found.

    echo.

    echo Checking Python Launcher...

    where py >nul 2>nul

    IF ERRORLEVEL 1 (

        echo.

        echo Error: Python installation could not be detected.

        echo Please install Python 3.10 or newer.

        echo Download:

        echo https://www.python.org/downloads/windows/

        SET SETUP_ERROR=1

        goto :ERROR_HANDLER

    ) ELSE (

        echo Python Launcher detected.

        SET PYTHON_BIN=py

    )

) ELSE (

    echo Python executable detected successfully.

    SET PYTHON_BIN=python

)


REM Mark Python as available.

SET PYTHON_FOUND=1


REM Display selected interpreter.

echo.

echo Selected Python command:

echo %PYTHON_BIN%

echo.


REM ============================================================
REM SECTION 5: VERIFY PYTHON VERSION
REM ============================================================

echo ============================================================
echo STEP 3: VERIFYING PYTHON VERSION
echo ============================================================

REM Display the installed Python version.

%PYTHON_BIN% --version

IF ERRORLEVEL 1 (

    echo Error: Unable to execute the selected Python interpreter.

    SET SETUP_ERROR=1

    goto :ERROR_HANDLER

)


REM Verify that Python 3.10 or newer is installed.

REM Python's sys module is used to compare version numbers.

%PYTHON_BIN% -c "import sys; exit(0 if sys.version_info >= (3,10) else 1)"

IF ERRORLEVEL 1 (

    echo.

    echo Error: Unsupported Python version detected.

    echo.

    echo This project requires Python 3.10 or newer.

    echo Please upgrade Python and run this script again.

    SET SETUP_ERROR=1

    goto :ERROR_HANDLER

)


echo.

echo Python version verification completed.

echo Compatible Python version detected.

echo.


REM ============================================================
REM SECTION 6: DISPLAY PYTHON INFORMATION
REM ============================================================

echo ============================================================
echo STEP 4: DISPLAYING PYTHON CONFIGURATION
echo ============================================================

REM Display the exact Python executable path.

echo Python executable location:

%PYTHON_BIN% -c "import sys; print(sys.executable)"

echo.


REM Display Python version details.

echo Python version details:

%PYTHON_BIN% -c "import sys; print(sys.version)"

echo.


REM Display Python installation prefix.

echo Python installation prefix:

%PYTHON_BIN% -c "import sys; print(sys.prefix)"

echo.


REM Verify pip availability.

%PYTHON_BIN% -m pip --version

IF ERRORLEVEL 1 (

    echo pip is not currently available.

    echo Attempting to initialize pip...

    %PYTHON_BIN% -m ensurepip --upgrade

    IF ERRORLEVEL 1 (

        echo Error: pip could not be initialized.

        SET SETUP_ERROR=1

        goto :ERROR_HANDLER

    )

)


echo Python configuration verified.

echo.


REM ============================================================
REM SECTION 7: NAVIGATE TO PROJECT ROOT
REM ============================================================

echo ============================================================
echo STEP 5: NAVIGATING TO PROJECT DIRECTORY
echo ============================================================

REM Change the current directory to the directory
REM containing this batch script.

cd /d "%~dp0"

IF ERRORLEVEL 1 (

    echo Error: Unable to access the project directory.

    SET SETUP_ERROR=1

    goto :ERROR_HANDLER

)


REM Display the current working directory.

echo Current project directory:

cd

echo.


REM Verify that the main application file exists.

IF NOT EXIST "%MAIN_FILE%" (

    echo Warning: %MAIN_FILE% was not found.

    echo The application may fail to launch.

    echo Please ensure that main.py exists in the project root.

    echo.

) ELSE (

    echo Main application file detected successfully.

)


REM Check the requirements file.

IF EXIST "%REQUIREMENTS_FILE%" (

    echo Requirements file detected.

) ELSE (

    echo Requirements file was not found.

    echo Base dependencies will be installed instead.

)


echo.

echo Project directory configuration completed.

echo.


REM ============================================================
REM SECTION 8: CHECK EXISTING VIRTUAL ENVIRONMENT
REM ============================================================

echo ============================================================
echo STEP 6: CHECKING VIRTUAL ENVIRONMENT
echo ============================================================

REM Check whether the virtual environment already exists.

IF EXIST "%VENV_DIR%\Scripts\python.exe" (

    echo Existing virtual environment detected.

    echo Reusing the existing environment.

) ELSE (

    echo No valid virtual environment was detected.

    echo A new virtual environment will be created.

)


REM ============================================================
REM SECTION 9: CREATE VIRTUAL ENVIRONMENT
REM ============================================================

echo ============================================================
echo STEP 7: CREATING PYTHON VIRTUAL ENVIRONMENT
echo ============================================================

REM Create the virtual environment only if necessary.

IF NOT EXIST "%VENV_DIR%\Scripts\python.exe" (

    echo.

    echo Creating virtual environment:

    echo %VENV_DIR%

    echo.

    %PYTHON_BIN% -m venv "%VENV_DIR%"

    IF ERRORLEVEL 1 (

        echo.

        echo Error: Virtual environment creation failed.

        SET SETUP_ERROR=1

        goto :ERROR_HANDLER

    )

    SET VENV_CREATED=1

    echo.

    echo Virtual environment created successfully.

) ELSE (

    echo.

    echo Existing virtual environment will be reused.

)


REM Verify the virtual environment Python executable.

IF NOT EXIST "%VENV_DIR%\Scripts\python.exe" (

    echo Error: Virtual environment Python executable is missing.

    SET SETUP_ERROR=1

    goto :ERROR_HANDLER

)


echo.

echo Virtual environment verification completed.

echo.


REM ============================================================
REM SECTION 10: ACTIVATE VIRTUAL ENVIRONMENT
REM ============================================================

echo ============================================================
echo STEP 8: ACTIVATING VIRTUAL ENVIRONMENT
echo ============================================================

REM Verify activation script availability.

IF NOT EXIST "%VENV_ACTIVATE%" (

    echo Error: Virtual environment activation script not found.

    SET SETUP_ERROR=1

    goto :ERROR_HANDLER

)


REM Activate the environment.

echo Activating %VENV_DIR%...

call "%VENV_ACTIVATE%"

IF ERRORLEVEL 1 (

    echo Error: Virtual environment activation failed.

    SET SETUP_ERROR=1

    goto :ERROR_HANDLER

)


REM Display the active environment.

echo.

echo Virtual environment activated successfully.

echo.


REM Display the Python executable currently being used.

echo Active Python executable:

python -c "import sys; print(sys.executable)"

echo.


REM Verify that the active Python interpreter belongs
REM to the project virtual environment.

python -c "import sys; print('Active environment:', sys.prefix)"

echo.


REM ============================================================
REM SECTION 11: VERIFY VIRTUAL ENVIRONMENT
REM ============================================================

echo ============================================================
echo STEP 9: VERIFYING ACTIVE ENVIRONMENT
echo ============================================================

REM Confirm that Python is running inside the virtual environment.

python -c "import sys; exit(0 if sys.prefix != sys.base_prefix else 1)"

IF ERRORLEVEL 1 (

    echo Error: Python is not running inside a virtual environment.

    SET SETUP_ERROR=1

    goto :ERROR_HANDLER

)


echo Active virtual environment verified.

echo.


REM ============================================================
REM SECTION 12: UPGRADE PIP
REM ============================================================

echo ============================================================
echo STEP 10: UPGRADING PIP
echo ============================================================

REM Upgrade pip to the latest available compatible version.

echo Checking current pip installation...

python -m pip --version

echo.

echo Upgrading pip...

python -m pip install --upgrade pip

IF ERRORLEVEL 1 (

    echo.

    echo Warning: pip upgrade encountered an error.

    echo Attempting to continue with the existing pip version.

) ELSE (

    echo.

    echo pip upgraded successfully.

)


REM Display the updated pip version.

echo.

echo Current pip version:

python -m pip --version

echo.


REM ============================================================
REM SECTION 13: INSTALL PROJECT DEPENDENCIES
REM ============================================================

echo ============================================================
echo STEP 11: INSTALLING PROJECT DEPENDENCIES
echo ============================================================

REM Check whether requirements.txt exists.

IF EXIST "%REQUIREMENTS_FILE%" (

    echo.

    echo Found requirements.txt.

    echo Installing all listed project dependencies.

    echo This process may take several minutes.

    echo.


    REM First installation attempt.

    python -m pip install -r "%REQUIREMENTS_FILE%"

    IF ERRORLEVEL 1 (

        echo.

        echo First dependency installation attempt failed.

        echo.

        echo Upgrading pip before retrying.

        python -m pip install --upgrade pip


        REM Second installation attempt.

        echo.

        echo Retrying dependency installation.

        python -m pip install -r "%REQUIREMENTS_FILE%"


        IF ERRORLEVEL 1 (

            echo.

            echo Error: Dependency installation failed.

            SET SETUP_ERROR=1

            goto :ERROR_HANDLER

        )

    )


    SET DEPENDENCIES_INSTALLED=1

    echo.

    echo Requirements installation completed successfully.


) ELSE (

    echo.

    echo No requirements.txt file was found.

    echo Installing base dependencies manually.

    echo.


    REM Install Streamlit.

    echo Installing Streamlit...

    python -m pip install streamlit

    IF ERRORLEVEL 1 (

        echo Error: Streamlit installation failed.

        SET SETUP_ERROR=1

        goto :ERROR_HANDLER

    )


    REM Install NumPy.

    echo.

    echo Installing NumPy...

    python -m pip install numpy

    IF ERRORLEVEL 1 (

        echo Error: NumPy installation failed.

        SET SETUP_ERROR=1

        goto :ERROR_HANDLER

    )


    REM Install Pandas.

    echo.

    echo Installing Pandas...

    python -m pip install pandas

    IF ERRORLEVEL 1 (

        echo Error: Pandas installation failed.

        SET SETUP_ERROR=1

        goto :ERROR_HANDLER

    )


    REM Install Matplotlib.

    echo.

    echo Installing Matplotlib...

    python -m pip install matplotlib

    IF ERRORLEVEL 1 (

        echo Error: Matplotlib installation failed.

        SET SETUP_ERROR=1

        goto :ERROR_HANDLER

    )


    SET DEPENDENCIES_INSTALLED=1

    echo.

    echo Base dependency installation completed.

)


REM ============================================================
REM SECTION 14: VERIFY STREAMLIT INSTALLATION
REM ============================================================

echo ============================================================
echo STEP 12: VERIFYING STREAMLIT INSTALLATION
echo ============================================================

REM Attempt to import Streamlit.

python -c "import streamlit"

IF ERRORLEVEL 1 (

    echo.

    echo Streamlit is not available in the environment.

    echo Attempting to install Streamlit.

    python -m pip install streamlit

    IF ERRORLEVEL 1 (

        echo.

        echo Error: Streamlit installation failed.

        SET SETUP_ERROR=1

        goto :ERROR_HANDLER

    )

)


REM Verify the Streamlit import again.

python -c "import streamlit; print('Streamlit version:', streamlit.__version__)"

IF ERRORLEVEL 1 (

    echo Error: Streamlit verification failed.

    SET SETUP_ERROR=1

    goto :ERROR_HANDLER

)


SET STREAMLIT_READY=1


echo.

echo Streamlit installation verified successfully.

echo.


REM ============================================================
REM SECTION 15: VERIFY NUMPY
REM ============================================================

echo ============================================================
echo STEP 13: VERIFYING NUMPY
echo ============================================================

REM Check NumPy import.

python -c "import numpy; print('NumPy version:', numpy.__version__)"

IF ERRORLEVEL 1 (

    echo Warning: NumPy could not be imported.

    echo Attempting installation.

    python -m pip install numpy

    IF ERRORLEVEL 1 (

        echo Error: NumPy installation failed.

        SET SETUP_ERROR=1

        goto :ERROR_HANDLER

    )

)


echo NumPy verification completed.

echo.


REM ============================================================
REM SECTION 16: VERIFY PANDAS
REM ============================================================

echo ============================================================
echo STEP 14: VERIFYING PANDAS
echo ============================================================

REM Check Pandas import.

python -c "import pandas; print('Pandas version:', pandas.__version__)"

IF ERRORLEVEL 1 (

    echo Warning: Pandas could not be imported.

    echo Attempting installation.

    python -m pip install pandas

    IF ERRORLEVEL 1 (

        echo Error: Pandas installation failed.

        SET SETUP_ERROR=1

        goto :ERROR_HANDLER

    )

)


echo Pandas verification completed.

echo.


REM ============================================================
REM SECTION 17: VERIFY MATPLOTLIB
REM ============================================================

echo ============================================================
echo STEP 15: VERIFYING MATPLOTLIB
echo ============================================================

REM Check Matplotlib import.

python -c "import matplotlib; print('Matplotlib version:', matplotlib.__version__)"

IF ERRORLEVEL 1 (

    echo Warning: Matplotlib could not be imported.

    echo Attempting installation.

    python -m pip install matplotlib

    IF ERRORLEVEL 1 (

        echo Error: Matplotlib installation failed.

        SET SETUP_ERROR=1

        goto :ERROR_HANDLER

    )

)


echo Matplotlib verification completed.

echo.


REM ============================================================
REM SECTION 18: VERIFY ALL REQUIRED PACKAGES
REM ============================================================

echo ============================================================
echo STEP 16: FINAL DEPENDENCY VERIFICATION
echo ============================================================

REM Verify all four main application dependencies together.

python -c "import streamlit, numpy, pandas, matplotlib; print('All base dependencies imported successfully.')"

IF ERRORLEVEL 1 (

    echo Error: One or more required packages are unavailable.

    SET SETUP_ERROR=1

    goto :ERROR_HANDLER

)


echo.

echo All base dependencies verified successfully.

echo.


REM ============================================================
REM SECTION 19: DISPLAY INSTALLED PACKAGE INFORMATION
REM ============================================================

echo ============================================================
echo STEP 17: DISPLAYING INSTALLED PACKAGES
echo ============================================================

REM Display the installed package list.

echo Installed packages:

python -m pip list

echo.


REM Display dependency information.

echo Dependency verification completed.

echo.


REM ============================================================
REM SECTION 20: CHECK STREAMLIT COMMAND
REM ============================================================

echo ============================================================
echo STEP 18: CHECKING STREAMLIT COMMAND
echo ============================================================

REM Check whether the Streamlit command is available.

where streamlit >nul 2>nul

IF ERRORLEVEL 1 (

    echo Streamlit executable was not found in PATH.

    echo.

    echo The application will use Python module execution instead.

    SET STREAMLIT_COMMAND_AVAILABLE=0

) ELSE (

    echo Streamlit executable detected successfully.

    SET STREAMLIT_COMMAND_AVAILABLE=1

)


REM ============================================================
REM SECTION 21: CHECK MAIN APPLICATION FILE
REM ============================================================

echo ============================================================
echo STEP 19: VERIFYING MAIN APPLICATION FILE
echo ============================================================

REM Confirm that main.py exists before launching.

IF NOT EXIST "%MAIN_FILE%" (

    echo.

    echo Error: Main application file was not found.

    echo Expected file:

    echo %MAIN_FILE%

    echo.

    echo Please ensure the project contains main.py.

    SET SETUP_ERROR=1

    goto :ERROR_HANDLER

)


echo Main application file verified.

echo.


REM ============================================================
REM SECTION 22: DISPLAY FINAL SETUP SUMMARY
REM ============================================================

echo ============================================================
echo STEP 20: SETUP SUMMARY
echo ============================================================

echo.

echo Application: %APP_NAME%

echo Project: %PROJECT_TITLE%

echo Python detected: %PYTHON_FOUND%

echo Virtual environment: %VENV_DIR%

echo Environment created during this run: %VENV_CREATED%

echo Dependencies installed: %DEPENDENCIES_INSTALLED%

echo Streamlit ready: %STREAMLIT_READY%

echo Main application: %MAIN_FILE%

echo.


REM ============================================================
REM SECTION 23: FINAL ENVIRONMENT CHECK
REM ============================================================

echo Performing final environment checks...

echo.


REM Check Python again.

python --version

IF ERRORLEVEL 1 (

    echo Error: Python environment verification failed.

    SET SETUP_ERROR=1

    goto :ERROR_HANDLER

)


REM Check pip again.

python -m pip --version

IF ERRORLEVEL 1 (

    echo Error: pip verification failed.

    SET SETUP_ERROR=1

    goto :ERROR_HANDLER

)


REM Check Streamlit again.

python -c "import streamlit"

IF ERRORLEVEL 1 (

    echo Error: Streamlit is unavailable.

    SET SETUP_ERROR=1

    goto :ERROR_HANDLER

)


REM Check application file again.

IF NOT EXIST "%MAIN_FILE%" (

    echo Error: main.py is missing.

    SET SETUP_ERROR=1

    goto :ERROR_HANDLER

)


echo.

echo All final environment checks passed.

echo.


REM ============================================================
REM SECTION 24: LAUNCH POKER AI SIMULATOR
REM ============================================================

echo ============================================================
echo STEP 21: LAUNCHING POKER AI SIMULATOR
echo ============================================================

echo.

echo Preparing Streamlit application...

echo.

echo Application:

echo %APP_NAME%

echo.

echo Entry point:

echo %MAIN_FILE%

echo.

echo The Streamlit UI should open in your browser shortly.

echo.

echo Press CTRL+C in this terminal to stop the application.

echo.

echo ============================================================

echo.

REM Give the user a moment to read the launch information.

timeout /t 2 /nobreak >nul


REM ============================================================
REM SECTION 25: START STREAMLIT APPLICATION
REM ============================================================

REM Use the Streamlit executable if it is available.
REM Otherwise use Python's module execution.

IF "!STREAMLIT_COMMAND_AVAILABLE!"=="1" (

    echo Starting application using Streamlit command...

    echo.

    streamlit run "%MAIN_FILE%"

) ELSE (

    echo Starting application using Python module...

    echo.

    python -m streamlit run "%MAIN_FILE%"

)


REM Capture application exit status.

SET APP_EXIT_CODE=!ERRORLEVEL!


REM ============================================================
REM SECTION 26: APPLICATION EXIT STATUS
REM ============================================================

echo.

echo ============================================================

echo APPLICATION EXECUTION STATUS

echo ============================================================

IF "!APP_EXIT_CODE!"=="0" (

    echo Poker AI Simulator exited successfully.

) ELSE (

    echo.

    echo Warning: Application exited with an error.

    echo Exit code: !APP_EXIT_CODE!

    echo.

    echo Please check the Streamlit terminal output.

    echo Verify that all project modules are installed.

    echo Check the main.py application entry point.

)


REM ============================================================
REM SECTION 27: CLEANUP
REM ============================================================

echo.

echo ============================================================

echo CLEANUP AND ENVIRONMENT SHUTDOWN

echo ============================================================

REM Deactivate the virtual environment if possible.

IF DEFINED VIRTUAL_ENV (

    echo Deactivating Python virtual environment...

    call deactivate

) ELSE (

    echo No active virtual environment detected.

)


REM Display completion information.

echo.

echo Setup process has finished.

echo.

echo Project directory:

echo %CD%

echo.

echo Virtual environment directory:

echo %VENV_DIR%

echo.

echo ============================================================

echo END OF POKER AI SIMULATOR SETUP

echo ============================================================

echo.


REM Return application exit status.

ENDLOCAL & EXIT /B %APP_EXIT_CODE%


REM ============================================================
REM SECTION 28: ERROR HANDLING
REM ============================================================

:ERROR_HANDLER

echo.

echo ============================================================

echo SETUP ERROR DETECTED

echo ============================================================

echo.

echo The setup process could not be completed.

echo.

echo Possible causes:

echo 1. Python is not installed correctly.

echo 2. Python is not available through PATH.

echo 3. The virtual environment could not be created.

echo 4. Dependency installation failed.

echo 5. Streamlit could not be installed.

echo 6. Required project files are missing.

echo 7. Network connectivity is unavailable.

echo.

echo Recommended actions:

echo.

echo - Verify your Python installation.

echo - Check your internet connection.

echo - Confirm requirements.txt is valid.

echo - Verify the project directory structure.

echo - Run the script again after resolving the issue.

echo.

echo ============================================================

echo SETUP TERMINATED

echo ============================================================

echo.

REM Attempt to deactivate the environment if activated.

IF DEFINED VIRTUAL_ENV (

    call deactivate

)


REM Exit with an error status.

ENDLOCAL

EXIT /B 1


REM ============================================================
REM END OF WINDOWS BATCH SCRIPT
REM ============================================================

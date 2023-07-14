# Download Python installer

$url = "https://www.python.org/ftp/python/3.11.4/python-3.11.4-amd64.exe"

# Specify where to save the installer
$output = "$env:TEMP\python-installer.exe"

# Download the installer
Invoke-WebRequest -Uri $url -OutFile $output

# Run the installer with silent options
Start-Process -FilePath $output -ArgumentList '/quiet InstallAllUsers=1 PrependPath=1' -Wait

# Install pip libraries
$pipLibraries = 'numpy', 'pandas', 'matplotlib', 'altgraph', 'certifi', 'cffi', 'charset-normalizer', 'cryptography', 'distlib', 'et-xmlfile', 'filelock', 'idna', 'jwt', 'msal', 'openpyxl', 'pefile', 'platformdirs', 'pycparser', 'python-dateutil', 'pytz', 'requests', 'sharepy', 'six', 'tzdata', 'xlrd', 'xlsxwriter', 'xlwt'

ForEach ($lib in $pipLibraries) {
    # Install the pip library
    pip install $lib
}

Write-Host "Installation completed!"
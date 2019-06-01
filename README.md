# Quicken
## Upon invoking the script
1. Copies a QDF file from the local repository to the working directory; if the file is present in the working directory you have the option of leaving it or replacing it with the file in the repository.
2. Invokes Quicken

## Upon exiting Quicken either
   Moves the QDF file in the working directory to the local    repository, replacing the original or

  * leaves the file in place in the working directory or
  * deletes the file in the working directory.

##2019-05-31 Major changes to LoadQuickenDb.ps1 v3.0.0
  * Added param statement
  * The script file is now invoked using powershell -noprofile -file thisscript -Filename theQuickenFile.QDF -speak

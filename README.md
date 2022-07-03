# Digital Satellite Data (Archive)
Program to extract private and unsecure info from the ISP Digital Satellite.


## Installation

Use the github cli [git-scm](https://git-scm.com/) to clone this Repo.

```bash
git clone https://github.com/Sspirax/Digital-Satellite-Data.git
```

Install the dependencies

```bash
pip install -r requirements.txt
```

## Usage

```bash
python run.py
```

Enter the required details.

Example:-
```bash
Enter the name of the building:- Waterlily
Enter total number of floors:- 18
Enter number of flats per floor:- 4
```

An xlsx file with the name of the building will be created in the root directory with all the user info.

## Vulnerability

Digital Satellite uses the same password ```123456``` for every user, and the username is the name of the building followed by the flat number.

This can be used to brute-force logins from each flat in a building and collect the info of every user who uses Digital Sattelite.

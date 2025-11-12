from Helper import RunningDinnerHelper

def get_teams(path):
    helper = RunningDinnerHelper(path)
    teams: list = []
    for team in helper.teams:
        teams.append(team[1])
    return teams


def open_pairings(teamName, path):
    helper = RunningDinnerHelper(path)
    pairings: list = []
    for check_team in helper.teams:
        if check_team[1] == teamName:
            for team in get_teams(path):
                if team not in check_team:
                    pairings.append(team)
            break
    return pairings

if __name__ == "__main__":
    team = input("Enter team name: ")
    pairings = open_pairings(team)
    print(f"Possible pairings for team '{team}':")
    for pairing in pairings:
        print(f"- {pairing}")
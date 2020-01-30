import json

class Player:
    """
    This class keeps track of players in smash league
    Args:
        name: username that player has chosen. Has to be unique, main identification for a player.
        wins: number of matches that a player has one
        losses: number of matches a player has lost
        last_five: list of the names that the other players. Only stores up to five names
        last five also enqueues players at the front and pops off the back, meaning the list
        is ordered like [newest, ... ,... ,... ,... , oldest]
    """
    def __init__(self, name = "", wins = 0, losses = 0,  last_five = []):
        self.name = name
        self.wins = wins
        self.losses = losses
        self.last_five = last_five
    
    def add_last_played(self , player_name):
        self.last_five.insert(0, player_name)
        if len(self.last_five) > 5:
            self.last_five.pop()
    
"""
Helper class to the player class. 
Converts a player object a json object 

example 
    p = Player(name = "Jeff" , wins = 15, losses= 7)
    p.last_five= ["Jeff1","Jeff2","Jeff3","Jeff4","Jeff5"]
    print(PlayerToJSON.PlayerToJSON(p))

expected ouput
{
  "name": "Jeff",
  "wins": 15,
  "losses": 7,
  "last_five": [
    "Jeff1",
    "Jeff2",
    "Jeff3",
    "Jeff4",
    "Jeff5"
  ]
}

"""

class PlayerToJSON:

    @staticmethod
    def PlayerToJSON(player):
        return json.dumps(player.__dict__) 
        


if __name__ == "__main__":
    p = Player(name = "Jeff" , wins = 15, losses= 7)
    p.last_five= ["Jeff1","Jeff2","Jeff3","Jeff4","Jeff5"]
    print(PlayerToJSON.PlayerToJSON(p))
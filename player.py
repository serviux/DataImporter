import json

class Player:
    
    def __init__(self, name = "", wins = 0, losses = 0,  last_five = []):
        self.name = name
        self.wins = wins
        self.losses = losses
        self.last_five = last_five
    
class PlayerToJSON:

    @staticmethod
    def PlayerToJSON(player):
        return json.dumps(player.__dict__)
        


if __name__ == "__main__":
    p = Player(name = "Jeff" , wins = 15, losses= 7)
    p.last_five= ["Jeff1","Jeff2","Jeff3","Jeff4","Jeff5"]
    print(PlayerToJSON.PlayerToJSON(p))
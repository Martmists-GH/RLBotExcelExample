import os
import itertools
import excel_parser as excel

from rlbot.agents.base_agent import BaseAgent, SimpleControllerState
from rlbot.utils.structures.game_data_struct import GameTickPacket


class ExcelAgent(BaseAgent):
    def initialize_agent(self, *args, **kwargs):
        fname = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bot.xlsx")
        self.model = excel.load_file(fname)
        self.chars = [f"{char}{num}" for char, num in
                      itertools.product("ABCDEFGH", [27])]
        
    def get_inputs(self, packet):
        ball_data = {
            "B2": packet.game_ball.physics.location.x,
            "B3": packet.game_ball.physics.location.y,
            "B4": packet.game_ball.physics.location.z,
            "B5": packet.game_ball.physics.velocity.x,
            "B6": packet.game_ball.physics.velocity.y,
            "B7": packet.game_ball.physics.velocity.z,
            "B8": packet.game_ball.physics.rotation.pitch,
            "B9": packet.game_ball.physics.rotation.roll,
            "B10": packet.game_ball.physics.rotation.yaw,
            "B11": packet.game_ball.physics.angular_velocity.x,
            "B12": packet.game_ball.physics.angular_velocity.y,
            "B13": packet.game_ball.physics.angular_velocity.z,
        }
        
        cars = [
            {
                chr(67+i) + "2": car.physics.location.x,
                chr(67+i) + "3": car.physics.location.y,
                chr(67+i) + "4": car.physics.location.z,
                chr(67+i) + "5": car.physics.velocity.x,
                chr(67+i) + "6": car.physics.velocity.y,
                chr(67+i) + "7": car.physics.velocity.z,
                chr(67+i) + "8": car.physics.rotation.pitch,
                chr(67+i) + "9": car.physics.rotation.roll,
                chr(67+i) + "10": car.physics.rotation.yaw,
                chr(67+i) + "11": car.physics.angular_velocity.x,
                chr(67+i) + "12": car.physics.angular_velocity.y,
                chr(67+i) + "13": car.physics.angular_velocity.z,
                chr(67+i) + "14": car.team,
                chr(67+i) + "15": int(car.jumped),
                chr(67+i) + "16": int(car.double_jumped),
                chr(67+i) + "17": car.boost
            }
            for i, car in enumerate(packet.game_cars) if i < 8
        ]
        
        other = {
            "A23": packet.num_cars,
            "B23": self.index
        }
        
        full = {}
        full.update(ball_data)
        for car in cars:
            full.update(car)
        full.update(other)
        
        return full
    
    def get_out(self):
        out = []
            
        for cellname in self.chars:
            try:
                out.append(float(self.model[cellname].evaluate()))
            except KeyError:
                raise Exception(f"Something led to cell {cellname} not being set!")
        
        return SimpleControllerState(*out)
        
    def get_output(self, packet: GameTickPacket) -> SimpleControllerState:
        inputs = self.get_inputs(packet)
        
        for k, v in inputs.items():
            self.model[k] = v
        
        print("Not opponent closer than car:", self.model["F42"].evaluate())
        print("Car in defensive:", self.model["G42"].evaluate())
        print("Res:", self.model["K33"].evaluate())
        print("Chosen action:", self.model["M33"].evaluate())
        
        return self.get_out()

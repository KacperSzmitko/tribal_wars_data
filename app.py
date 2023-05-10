from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver import ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import json
import os
from typing import List
from dataclasses import dataclass, field
from selenium.webdriver.support.ui import Select
import xlsxwriter


URL = "https://www.plemiona.pl/"
LOGGED = "menu_row"
TRIBE_XPATH = "/html/body/table/tbody/tr[1]/td[2]/div/table/tbody/tr/td/table/tbody/tr/td[10]"
TRIBE_MEMBERS = "/html/body/table/tbody/tr[2]/td[2]/table[3]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[4]"
TRIBE_ARMY = '//*[@id="ally_content"]/table/tbody/tr/td[2]/a'
TRIBE_BUILDINGS = '//*[@id="ally_content"]/table/tbody/tr/td[3]/a'
TRIBE_DEFF = '//*[@id="ally_content"]/table/tbody/tr/td[4]/a'
PLAIN_TABLE = '//*[@id="ally_content"]/div/div/table'


@dataclass
class Options:
    username: str
    password: str
    army: bool = False
    buildings: bool = True
    deff: bool = False


@dataclass
class Building:
    nazwa: str = ''
    punkty: int = 0
    ratusz: int = 0
    koszary: int = 0
    stajnia: int = 0
    warsztat: int = 0
    pałac: int = 0
    kuźnia: int = 0
    plac: int = 0
    piedestał: int = 0
    rynek: int = 0
    tartak: int = 0
    cegielnia: int = 0
    huta: int = 0
    zagroda: int = 0
    spichlerz: int = 0
    mur: int = 0


@dataclass
class Army:
    nazwa: str = ''
    piki: int = 0
    miecze: int = 0
    topory: int = 0
    zwiad: int = 0
    lk: int = 0
    ck: int = 0
    tarany: int = 0
    katsy: int = 0
    rycek: int = 0
    grube: int = 0
    komendy: int = 0
    ataki: int = 0


@dataclass
class Player:
    name: str
    buildings: List[Building] = field(default_factory=list)
    army: List[Army] = field(default_factory=list)
    deff: List[Building] = field(default_factory=list)

    def __eq__(self, __value: object) -> bool:
        return self.name == __value


@dataclass
class Config:
    options: Options
    build_requirements: Building
    army_requirements: Army

    def __post_init__(self):
        if isinstance(self.options, dict):
            self.options = Options(**self.options)
        if isinstance(self.build_requirements, dict):
            self.build_requirements = Building(**self.build_requirements)
        if isinstance(self.army_requirements, dict):
            self.army_requirements = Army(**self.army_requirements)


class DataBot:

    def __init__(self) -> None:
        self.workbook = xlsxwriter.Workbook('data.xlsx')
        self.config: Config = None
        self.load_options()
        opts = ChromeOptions()
        opts.add_argument("start-maximized")
        opts.add_experimental_option("detach", True)
        opts.add_experimental_option("useAutomationExtension", False)
        opts.add_experimental_option('excludeSwitches', ['enable-automation'])
        self.browser = webdriver.Chrome(
            ChromeDriverManager().install(), options=opts)
        self.login()
        self.players: List[Player] = []
        self.first = True
        self.PASSED_FORMAT = self.workbook.add_format()
        self.PASSED_FORMAT.set_bg_color('green')
        self.FAILED_FORMAT = self.workbook.add_format()
        self.FAILED_FORMAT.set_bg_color('red')

    def load_options(self) -> dict:
        if not os.path.exists("config.json"):
            self.make_default_options()
        else:
            with open("config.json", "r+", encoding='utf-8') as file:
                try:
                    data = json.load(file)
                    self.config = Config(**data)
                except (json.decoder.JSONDecodeError):
                    self.make_default_options()

    def make_default_options(self):
        with open("config.json", "w+", encoding='utf-8') as file:
            username = input("Please enter your login: ")
            password = input("Please enter your password: ")
            army = input("Do you want army info? (y/n)") in ('y', 'Y', '')
            buildings = input(
                "Do you want buildings info? (y/n)") in ('y', 'Y', '')
            deff = input("Do you want deff info? (y/n)") in ('y', 'Y', '')
            options = Options(
                username, password, army, buildings, deff)
            self.config = Config(options, Building(), Army())

            def default(o):
                o.__dict__.pop('nazwa', '')
                o.__dict__.pop('punkty', '')
                return o.__dict__
            file.write(json.dumps(self.config.__dict__,
                       default=default, indent=4))

    def login(self):
        self.browser.get(URL)
        self.browser.find_element(
            By.NAME, "username").send_keys(self.config.options.username)
        self.browser.find_element(By.NAME, "password").send_keys(
            self.config.options.password)
        self.browser.find_element(By.CLASS_NAME, "btn-login").click()
        WebDriverWait(self.browser, 60).until(
            EC.presence_of_element_located((By.ID, LOGGED)))

    def parse_to_int(self, element: str) -> int:
        try:
            return int(element.text)
        except ValueError:
            return element.text

    def get_player(self, name: str):
        if not self.first:
            player = next((x for x in self.players if x == name), None)
        else:
            player = Player(name)
            self.players.append(player)
        return player

    def get_data(self, type: str):
        select = Select(
            self.browser.find_element(By.NAME, "player_id"))
        options = select.options
        for index in range(1, len(options) - 1):
            select.select_by_index(index)
            select = Select(
                self.browser.find_element(By.NAME, "player_id"))
            player_name = select.first_selected_option.text
            player = self.get_player(player_name)
            table_id = self.browser.find_element(
                By.XPATH, PLAIN_TABLE)
            rows = table_id.find_elements(By.TAG_NAME, "tr")[1:]
            for row in rows:
                cols = map(self.parse_to_int,
                           row.find_elements(By.TAG_NAME, "td"))
                if type == "buildings":
                    player.buildings.append(Building(*cols))
                elif type == "army":
                    player.army.append(Army(*cols))

    def make_sheet(self, type: str, validate: bool = False):
        if type == "buildings":
            sheet = self.workbook.add_worksheet("Budynki")
            template = Building.__match_args__
            requirements = self.config.build_requirements
        elif type == "army":
            sheet = self.workbook.add_worksheet("Wojsko")
            template = Army.__match_args__
            requirements = self.config.army_requirements

        col = 'B'
        for k in template:
            if k != "punkty":
                sheet.write(f'{col}1', k)
                col = chr(ord(col) + 1)

        col, row = 'B', 2
        for player in self.players:
            cell_format = self.workbook.add_format(
                {'bold': True})
            start_row, failed = row, False
            for entity in getattr(player, type):
                for i, k in enumerate(entity.__dict__):
                    if k not in ('punkty'):
                        value = entity.__dict__[k]
                        if value == "?":
                            sheet.write(f'{col}{row}', value,
                                        self.FAILED_FORMAT)
                            failed = True
                            continue
                        if validate:
                            expected_value = getattr(requirements, k)
                            if isinstance(expected_value, int) and value < expected_value:
                                # Color cell with failed thing
                                sheet.write(
                                    f'{col}{row}', value, self.FAILED_FORMAT)
                                failed = True
                            else:
                                sheet.write(f'{col}{row}', value)
                        else:
                            sheet.write(f'{col}{row}', value)
                        col = chr(ord(col) + 1)
                row += 1
                col = 'B'
            # Color player name if failed
            if failed and validate:
                cell_format.set_bg_color('red')
            sheet.write(f'A{start_row}', player.name, cell_format)

    def run(self):
        self.browser.find_element(By.XPATH, TRIBE_XPATH).click()
        self.browser.find_element(By.XPATH, TRIBE_MEMBERS).click()
        if self.config.options.army:
            self.browser.find_element(By.XPATH, TRIBE_ARMY).click()
            self.get_data("army")
            self.make_sheet("army", validate=True)
            self.first = False
        if self.config.options.buildings:
            self.browser.find_element(By.XPATH, TRIBE_BUILDINGS).click()
            self.get_data("buildings")
            self.make_sheet("buildings", validate=True)
            self.first = False
        self.workbook.close()
        self.browser.quit()
        # with open("test_data.json", "w+") as file:
        #     file.write(json.dumps(
        #         self.players, default=lambda x: x.__dict__, indent=4))


if __name__ == "__main__":
    DataBot().run()

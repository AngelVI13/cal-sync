import json
import time
import unittest
import subprocess
import configparser
from pathlib import Path
from typing import Set, Dict
from dataclasses import dataclass

from selenium import webdriver


@dataclass
class MeetingInfo:
    from_: str
    when: str
    sent_time: str  # when the meeting invite was sent

    def __eq__(self, other) -> bool:
        return hash(self) == hash(other)

    def __hash__(self) -> int:
        return hash((self.from_, self.when, self.sent_time))

    def to_dict(self) -> Dict[str, str]:
        return self.__dict__

    @classmethod
    def from_dict(cls, d: Dict[str, str]):
        return cls(from_=d["from_"], when=d["when"], sent_time=d["sent_time"])


class SimpleCalculatorTests(unittest.TestCase):
    config_path = Path("config.ini")

    @classmethod
    def setUpClass(self):
        if not (self.config_path.exists() and self.config_path.is_file()):
            raise Exception(
                f"Failed to find config file: {self.config_path.absolute().as_posix()}"
            )

        self.config = configparser.ConfigParser()
        self.config.read(self.config_path)

        settings = self.config["settings"]
        self.run_driver = settings.getboolean("RunWinAppDriver")
        self.meeting_folder_name = settings["MeetingsFolderName"]
        self.forward_address = settings["ForwardAddress"]

        self.history_file = Path(settings["HistoryFile"])
        self.already_sent: Dict[int, MeetingInfo] = {}
        if self.history_file.exists() and self.history_file.is_file():
            s = self.history_file.read_text(encoding="utf-8")
            data = json.loads(s)
            self.already_sent = {}
            for info in data:
                d = MeetingInfo.from_dict(info)
                self.already_sent[hash(d)] = d

        if self.run_driver:
            self.win_app_driver = subprocess.Popen(
                args=[settings["WinAppDriverExePath"]],
                shell=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )

        desired_caps = {}
        desired_caps["app"] = settings["OutlookExePath"]
        self.driver = webdriver.Remote(
            command_executor="http://127.0.0.1:4723", desired_capabilities=desired_caps
        )
        time.sleep(2)

    @classmethod
    def tearDownClass(self):
        self.driver.quit()
        if self.run_driver:
            self.win_app_driver.terminate()

    def openMeetingRequests(self):
        meetingRequestsFolder = self.driver.find_element_by_name(
            self.meeting_folder_name
        )
        meetingRequestsFolder.click()
        time.sleep(1)

    def handleSeriesPopup(self, prev_handles: Set[str]):
        """Handle popup that asks if to send only this occurance or series.
        This only happens for recurring meetings.
        """
        # NOTE: if a popup (REMINDER or sth else) appears before this part then
        # the script won't be able to find the correct popup
        new_handles = set(self.driver.window_handles)
        handle_diff = new_handles - prev_handles

        if len(handle_diff) == 0:
            return

        popup_handle, *_ = handle_diff
        self.driver.switch_to.window(popup_handle)

        try:
            self.driver.find_element_by_name("Send the series.")
        except:
            pass
        else:
            # a popup is shown -> accept it
            self.driver.find_element_by_name("OK").click()
            time.sleep(1)

    def forwardMail(self):
        # forward email
        msg_actions = self.driver.find_element_by_name("Message Actions")
        other_actions = msg_actions.find_element_by_name("Other Actions")
        other_actions.click()
        time.sleep(1)

        prev_handles = set(self.driver.window_handles)
        main_window = self.driver.current_window_handle

        self.driver.find_element_by_name("Forward").click()
        time.sleep(1)

        self.handleSeriesPopup(prev_handles)

        new_handles = set(self.driver.window_handles)
        handle_diff = new_handles - prev_handles
        if len(handle_diff) == 0:
            # In cases where "Forward" button is disabled - no new window openes
            return False

        mail_forward, *_ = handle_diff
        self.driver.switch_to.window(mail_forward)

        to_box = self.driver.find_element_by_class_name("RichEdit20WPT")
        to_box.send_keys(self.forward_address)

        self.driver.find_element_by_name("Send").click()
        self.driver.switch_to.window(main_window)
        return True

    def mailInfo(self, m) -> MeetingInfo:
        m.click()

        time.sleep(1)  # wait for elements to load

        when = self.driver.find_elements_by_class_name(
            "Change Highlighting Edit Control"
        )[0].text
        from_ = self.driver.find_element_by_name("From").text
        sent = self.driver.find_element_by_name("Sent").text

        new_info = MeetingInfo(from_=from_, when=when, sent_time=sent)
        return new_info

    def syncToFile(self, info: Dict[int, MeetingInfo]):
        data = [v.to_dict() for v in info.values()]
        self.history_file.write_text(json.dumps(data), encoding="utf-8")

    def test_initialize(self):
        self.driver.maximize_window()
        try:
            maximizeFolder = self.driver.find_element_by_name("Folder Pane Minimized")
        except:
            pass
        else:
            maximizeFolder.click()

        self.openMeetingRequests()

        email_box = self.driver.find_element_by_name("Table View")
        try:
            down_btn = email_box.find_element_by_name("Line down")
        except:
            # if we can't find the down button it means that there is not a lot of
            # emails -> no need to scroll down
            down_btn = None
            done = True
        else:
            done = False

        info = self.already_sent
        for _ in range(20):
            meetings = email_box.find_elements_by_class_name("LeafRow")
            if not meetings:
                break
            # set_trace()

            processed = 0
            for m in meetings:
                if m.location["x"] < 100 and m.location["y"] < 100:
                    continue

                processed += 1

                new_info = self.mailInfo(m)
                if hash(new_info) in info:
                    # the first time you encounter a duplicate -> flag to stop
                    # but keep going to make sure all potential meetings are clicked
                    # This solves the issue when you cant scroll to a full set of
                    # new meetings so the view still contains already processed meetings
                    # and a few new ones at the bottom
                    done = True
                    continue

                success = self.forwardMail()
                if not success:
                    print(f"Failed to forward email (forwarding disable): {new_info}")

                info[hash(new_info)] = new_info
                self.syncToFile(info)
                # set_trace()

            if done or down_btn is None:
                break

            # scroll down the page. each item ~= 1 scroll
            for _ in range(processed):
                down_btn.click()


if __name__ == "__main__":
    suite = unittest.TestLoader().loadTestsFromTestCase(SimpleCalculatorTests)
    unittest.TextTestRunner(verbosity=2).run(suite)

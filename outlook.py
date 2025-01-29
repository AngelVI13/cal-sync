import time
import unittest
import subprocess
from typing import Set
from pprint import pprint
from dataclasses import dataclass
from ipdb import set_trace
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


class SimpleCalculatorTests(unittest.TestCase):
    run_driver = False

    @classmethod

    def setUpClass(self):
        # TODO: update this script to automatically start the winappdriver process!!
        #set up appium
        if self.run_driver:
            self.win_app_driver = subprocess.Popen(
                args=[r"C:\Program Files (x86)\Windows Application Driver\WinAppDriver.exe"],
                shell=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )

        desired_caps = {}
        desired_caps["app"] = r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
        self.driver = webdriver.Remote(
            command_executor='http://127.0.0.1:4723',
            desired_capabilities= desired_caps)
        time.sleep(2)

    @classmethod
    def tearDownClass(self):
        self.driver.quit()
        if self.run_driver:
            self.win_app_driver.terminate()

    def openMeetingRequests(self):
        meetingRequestsFolder = self.driver.find_element_by_name("Meeting Requests")
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
            time.sleep(0.5)

    def forwardMail(self):
        # forward email
        msg_actions = self.driver.find_element_by_name("Message Actions")
        other_actions = msg_actions.find_element_by_name("Other Actions")
        other_actions.click()
        time.sleep(0.5)

        prev_handles = set(self.driver.window_handles)
        main_window = self.driver.current_window_handle

        self.driver.find_element_by_name("Forward").click()
        time.sleep(0.5)

        self.handleSeriesPopup(prev_handles)

        new_handles = set(self.driver.window_handles)
        mail_forward, *_ = new_handles - prev_handles
        self.driver.switch_to.window(mail_forward)

        to_box = self.driver.find_element_by_name("RichEdit20WPT")
        to_box.send_keys("angel.iliev@qdevtechnologies.com")

        self.driver.find_element_by_name("Send").click()
        self.driver.switch_to.window(main_window)

    def mailInfo(self, m) -> MeetingInfo:
        m.click()

        time.sleep(0.5)  # wait for elements to load

        when = self.driver.find_elements_by_class_name("Change Highlighting Edit Control")[0].text
        from_ = self.driver.find_element_by_name("From").text
        sent = self.driver.find_element_by_name("Sent").text

        new_info = MeetingInfo(from_=from_, when=when, sent_time=sent)
        return new_info


    def test_initialize(self):
        self.driver.maximize_window()
        # TODO: for this to work you have to create a folder 'Meeting Requests' and 
        # create a rule to move/copy all incoming meeting requests to that folder
        # then this script will look through that folder and forward those requests to an
        # email of your choice
        try:
            maximizeFolder = self.driver.find_element_by_name("Folder Pane Minimized")
        except:
            pass
        else:
            maximizeFolder.click()

        # TODO: also this requires disabling date headers in outlook for Meeting Requests folder
        #  for easier processing of meeting requests: 
        #  https://answers.microsoft.com/en-us/outlook_com/forum/all/how-do-i-remove-the-date-grouping-in-the-new/e3267590-6abd-4545-b8c4-ddf9317dbbd7

        self.openMeetingRequests()

        # TODO: figure out how to click all of them (including scrolling down)
        # because this just finds all visible elements
        email_box = self.driver.find_element_by_name("Table View")
        down_btn = email_box.find_element_by_name("Line down")

        info = []
        done = False
        for _ in range(20):
            # TODO: can't use the set directly cause the elements are based on position so when you scrol
            # the same email is considered a completely different element
            meetings = email_box.find_elements_by_class_name("LeafRow")
            # set_trace()

            processed = 0
            for m in meetings:
                if m.location["x"] < 100 and m.location["y"] < 100:
                    continue

                processed += 1

                new_info = self.mailInfo(m)
                if new_info in set(info):
                    # the first time you encounter a duplicate -> flag to stop
                    # but keep going to make sure all potential meetings are clicked
                    # This solves the issue when you cant scroll to a full set of 
                    # new meetings so the view still contains already processed meetings 
                    # and a few new ones at the bottom
                    done = True
                    continue

                self.forwardMail()
                info.append(new_info)
                set_trace()

            if done:
                break

            # scroll down the page. each item ~= 1 scroll
            for _ in range(processed):
                down_btn.click()

        pprint(info)
        print(len(info))
        set_trace()


if __name__ == '__main__':
    suite = unittest.TestLoader().loadTestsFromTestCase(SimpleCalculatorTests)
    unittest.TextTestRunner(verbosity=2).run(suite)

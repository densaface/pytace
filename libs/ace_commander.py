#!/usr/bin/env python
# -*- coding: utf-8 -*-

import subprocess
import os
import shlex
import time
import win32pipe, win32file
import re
import pywintypes
import win32event
import winerror

debug_level = 1
use_overlapped = 0

import win32gui
import re


class WindowMgr:
    """Encapsulates some calls to the winapi for window management"""

    def __init__ (self):
        """Constructor"""
        self._handle = None

    def find_window(self, class_name, window_name=None):
        """find a window by its class_name"""
        self._handle = win32gui.FindWindow(class_name, window_name)

    def _window_enum_callback(self, hwnd, wildcard):
        """Pass to win32gui.EnumWindows() to check all the opened windows"""
        if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) is not None:
            self._handle = hwnd

    def find_window_wildcard(self, wildcard):
        """find a window whose title matches the wildcard regex"""
        self._handle = None
        win32gui.EnumWindows(self._window_enum_callback, wildcard)

    def set_foreground(self):
        """put the window in the foreground"""
        win32gui.SetForegroundWindow(self._handle)

class AceComs():
    def __init__(self, log_file = ''):
        self.log_file = log_file
    def pipeReq(self, mes, attempts = 60):
        if debug_level >= 3:
            print (mes)
        for ii in range(attempts):
            try:
                pipe_req = os.open("\\\\.\\pipe\\ace_exchange_%d" % 369, os.O_RDWR)
                break
            except:
                if debug_level >= 3:
                    print ("new attempt to create req pipe")
                time.sleep(1)
        try:
            pipe_req
        except Exception as e:
            print("check if ACE started!")
            print(e)
            return False
        os.write(pipe_req, mes.encode())
        os.close(pipe_req)
        return pipe_req
    def createAnswerPipe(self):
        try:
            p = win32pipe.CreateNamedPipe(r'\\.\pipe\ace_answer_%d' % 369,
                                          win32pipe.PIPE_ACCESS_DUPLEX,
                                          win32pipe.PIPE_TYPE_MESSAGE | win32pipe.PIPE_WAIT,
                                          1, 65536, 65536, 300, None)
        except Exception as e:
            print(e)
            import ipdb;
            ipdb.set_trace()
            p = win32pipe.CreateNamedPipe(r'\\.\pipe\ace_answer_%d' % 369,
                                          win32pipe.PIPE_ACCESS_DUPLEX,
                                          win32pipe.PIPE_TYPE_MESSAGE | win32pipe.PIPE_WAIT,
                                          1, 65536, 65536, 300, None)
        return p

    def getAnswer(self, pipe):
        for ii in range(10):
            try:
                if not use_overlapped:
                    rc = win32pipe.ConnectNamedPipe(pipe, None)
                else:
                    overlapped = pywintypes.OVERLAPPED()
                    overlapped.hEvent = win32event.CreateEvent(None, 0, 0, None)
                    rc = win32pipe.ConnectNamedPipe(pipe, overlapped)
                    if rc == winerror.ERROR_PIPE_CONNECTED:
                        win32event.SetEvent(overlapped.hEvent)
                    rc = win32event.WaitForSingleObject(overlapped.hEvent, 10000)
                break
            except Exception as e:
                print(e)
                if ii > 2:
                    import ipdb; ipdb.set_trace()
                return bytes("ConnectNamedPipe error", 'cp1251')
                # try:
                #     win32pipe.DisconnectNamedPipe(pipe)
                # except Exception as e:
                #     print(e)
                # time.sleep(0.5)
                # print("before getAnswer 2")
                # # p.close()
                # # del p
                # # p = self.createAnswerPipe()
                # return self.getAnswer(pipe, attempt + 1)
        full_ans = win32file.ReadFile(pipe, 4096)
        win32pipe.DisconnectNamedPipe(pipe)
        pipe.close()
        del pipe
        return full_ans[1]

    def SaveDesktop(self, out_file, l,t,r,b):
        args = []
        args.append('-save_image')
        # args.append('-main_path=\'%s\'' % script_path)
        args.append('-out=\'%s\'' % out_file)
        args.append('-ltrb=%d,%d,%d,%d' % (l,t,r,b))

        p = self.createAnswerPipe()
        pipe_descr = self.pipeReq(' '.join(args))
        answer = self.getAnswer(p).decode('cp1251')
    
        self.log_print(time.asctime( time.localtime( time.time() ) ) + (" : SaveDesktop (%s) : " % out_file) + answer[0:-1])
        if answer.find("ok") > -1:
            return True
        else:
            return False

    def findImage(self, out_file, l, t, r, b):
        args = []
        # args.append('-save_image')
        args.append('-main_path=\'%s\'' % script_path)
        args.append('-out=\'%s\'' % out_file)
        args.append('-ltrb=%d,%d,%d,%d' % (l, t, r, b))

        p = self.createAnswerPipe()
        pipe_descr = self.pipeReq(' '.join(args))
        answer = self.getAnswer(p).decode('cp1251')

        print(time.asctime(time.localtime(time.time())) + " : SaveDesktop : " + answer[0:-1])
        if answer.find("ok") > -1:
            return True
        else:
            return False

    # find_picts - массив изображений для поиска, может быть и строчной переменной для задания единственно изображения
    # prefer_x   - массив координат, где картинка ожидается быть найденной (для ускорения поиска)
    # max_radius - максимальный радиус поиска относительно предпочтительных координат
    # deviation  - разрешенное отклонение картинки в цветах на пиксель
    def GetImageCoord(self, src_file, find_picts, main_path, attempt=1,
                      print_out=True):
        # print "GetImageCoord:"
        # print find_picts
        args = ['-findpict', ]
        img_list = [] # for report
        if type(find_picts) is list:
            for ii in range(len(find_picts)):
                args.append('-find_pict%d=\'%s\'' % (ii + 1, find_picts[ii]['img']))
                fi_cf = find_picts[ii]['img'].rfind('/')
                if fi_cf > -1:
                    img_list.append(find_picts[ii]['img'][fi_cf + 1:])
                else:
                    img_list.append(find_picts[ii])
                if 'prefer_x' in find_picts[ii]:
                    args.append('-prefer_x%d=\'%d\'' % (ii + 1, find_picts[ii]['prefer_x']))
                if 'prefer_y' in find_picts[ii]:
                    args.append('-prefer_y%d=\'%d\'' % (ii + 1, find_picts[ii]['prefer_y']))
                if 'max_radius' in find_picts[ii]:
                    args.append('-max_radius%d=\'%d\'' % (ii + 1, find_picts[ii]['max_radius']))
                if 'max_radius_y' in find_picts[ii]:
                    args.append('-max_radius_y%d=\'%d\'' % (ii + 1, find_picts[ii]['max_radius_y']))
                if 'deviation' in find_picts[ii]:
                    args.append('-deviation%d=\'%d\'' % (ii + 1, find_picts[ii]['deviation']))
                if 'blur' in find_picts[ii]:
                    args.append('-blur%d=\'%d\'' % (ii + 1, find_picts[ii]['blur']))
        elif type(find_picts) is dict:
            args.append('-find_pict%d=\'%s\'' % (1, find_picts['img']))
            fi_cf = find_picts['img'].rfind('/')
            if fi_cf > -1:
                img_list.append(find_picts['img'][fi_cf+1:])
            else:
                img_list.append(find_picts['img'])
            if 'prefer_x' in find_picts:
                args.append('-prefer_x%d=\'%d\'' % (1, find_picts['prefer_x']))
            if 'prefer_y' in find_picts:
                args.append('-prefer_y%d=\'%d\'' % (1, find_picts['prefer_y']))
            if 'max_radius' in find_picts:
                args.append('-max_radius%d=\'%d\'' % (1, find_picts['max_radius']))
            if 'max_radius_y' in find_picts:
                args.append('-max_radius_y%d=\'%d\'' % (1, find_picts['max_radius_y']))
            if 'deviation' in find_picts:
                args.append('-deviation%d=\'%d\'' % (1, find_picts['deviation']))
            if 'blur' in find_picts:
                args.append('-blur%d=\'%d\'' % (1, find_picts['blur']))
        else:
            args.append('-find_pict%d=\'%s\'' % (1, find_picts))

        args.append('-main_path=\'%s\'' % main_path)
        args.append('-src=\'%s\'' % src_file)

        # print ' '.join(args)
        time1 = time.time()
        p = self.createAnswerPipe()
        pipe_descr = self.pipeReq(' '.join(args))
        answer = self.getAnswer(p).decode('cp1251')

        # координаты нахождения картинки
        xx = -1
        yy = -1
        fi1 = answer.find("coordinates = ")
        fi2 = answer.find(":")
        if fi1 > -1 and fi2 > -1:
            xx = answer[fi1 + 14:fi2]
            yy = answer[fi2 + 1:answer.find(', deviation')]
        elif answer.find("Not found any image") > -1:
            if len(img_list) == 1:
                answer = answer.replace("any image", "image " + img_list[0])
                print(answer)
            else:
                print(answer + ", img_list = " + str(img_list))
            return -1, "not found image", -1, -1, -1
        deviat = answer[answer.find(', deviation') + 14: answer.find(' color ')]

        # какая по счету картинка найдена
        num_image = -1
        fi1 = answer.find("Found image #")
        fi2 = answer.find(", coordinates")
        if fi1 > -1 and fi2 > -1:
            num_image = int(answer[fi1 + 13:fi2])
            if type(find_picts) is list:
                if print_out:
                    print("FOUND IMG = %s (%d:%d), deviation = %s" % (\
                        find_picts[num_image - 1]['img'], int(xx), int(yy), deviat))
            elif type(find_picts) is dict:
                if print_out:
                    print("FOUND IMG = %s (%d:%d), deviation = %s" % (\
                        find_picts['img'], int(xx), int(yy), deviat))
        if int(xx) < 0:
            print(answer)
            import ipdb; ipdb.set_trace()
            return self.GetImageCoord(src_file, find_picts, main_path, attempt+1,True)
            # raise Exception("xx < 0")

        return 0, "success", int(xx), int(yy), int(num_image)

    def waitReplay(self, attempts=1200):
        for ii in range(attempts):
            result_replay = self.getShortLog()
            m = re.search("ReplayNTimes\(N=\d+\):[\w\s]+[Ss]uccess", result_replay)
            if m:
                # examples success result_replay:
                # "ReplayNTimes(N=1): Success"
                # "ReplayNTimes(N=1): stopped success by script on act 2 (01:01:32)"
                print(time.asctime(time.localtime(time.time())) + ": replayed successfull")
                print(result_replay)
                return True, result_replay
            if result_replay.find(u'topped by user') > -1:
                print(time.asctime(time.localtime(time.time())) + ": Replay stopped by user!")
                print(result_replay)
                return False, result_replay
            else:
                # print(time.asctime(time.localtime(time.time())))
                # print(result_replay)
                time.sleep(1)
        print (time.asctime(time.localtime(time.time())) + ": Didn't wait the end of replay")
        return False, result_replay

    def stopReplay(self):
        args = []
        args.append('-stop')

        p = self.createAnswerPipe()
        self.pipeReq(' '.join(args))
        answer = self.getAnswer(p).decode('cp1251')
        print(time.asctime(time.localtime(time.time())) + " : StopReplay : " + answer[0: -1])
        time.sleep(2)
        if answer.find("ok") > -1:
            return True
        else:
            return False

    def Replay(self, aip='', wait_sec = 1200):
        if aip:
            res_load = self.loadAip(aip)
            time.sleep(1)
            if not res_load:
                import ipdb; ipdb.set_trace()
                return False, "loadAip error"
        args = []
        args.append('-play')
    
        p = self.createAnswerPipe()
        self.pipeReq(' '.join(args))
        answer = self.getAnswer(p).decode('cp1251')
        print (time.asctime( time.localtime( time.time() ) ) + " : Replay start : " + answer[0: -1])
        if answer.find("ok") > -1:
            if aip == '':
                return True, answer
        else:
            if aip == '':
                return False, answer
            else:
                print(answer)
                raise Exception("Wrong replay start")
        if aip:
            time.sleep(1)
            return self.waitReplay(wait_sec)
    
    def loadAip(self, aipFile, attempt=1):
        args = []
        args.append('-load')
        args.append(aipFile)
        p = self.createAnswerPipe()
        self.pipeReq(' '.join(args))
        answer = self.getAnswer(p).decode('cp1251')
        print (time.asctime( time.localtime( time.time() ) ) + " : loadAip : " + answer[0: -1])
        if answer.find("error_load") > -1 and len(answer) < 15:
            if attempt > 5:
                print("aipFile=" + aipFile)
                return False
            time.sleep(1)
            import ipdb; ipdb.set_trace()
            return self.loadAip(aipFile, attempt+1)
        else:
            return True

    def setClip(self, str):
        command = 'echo | set /p nul=' + str.strip() + '| clip'
        # command = 'echo ' + str.strip() + '| clip'
        os.system(command)

        # r = tk.Tk()
        # r.withdraw()
        # r.clipboard_clear()
        # r.clipboard_append(str)
        # r.update()  # now it stays on the clipboard after the window is closed
        # r.destroy()

    def setVar(self, var, num_attemt=1):
        args = []
        args.append('-setvar')
        args.append(var)
        p = self.createAnswerPipe()
        pipe_descr = self.pipeReq(' '.join(args))
        answer = self.getAnswer(p).decode('cp1251')
        print (time.asctime( time.localtime( time.time() ) ) + " : setVar (%s): " % var + answer[0: -1])
        if answer.find("ok") > -1:
            return True
        else:
            if num_attemt < 5:
                print ("setVar %s warning, new attempt", var)
                time.sleep(200)
                return self.setVar(var, num_attemt+1)
            raise Exception("setVar %s error", var)

    def clearReplayScreenshots(self, path_name):
        def clear_dir(dirc):
            for (dirpath, dirnames, filenames) in os.walk(dirc):
                for file in filenames:
                    os.remove(dirc + "\\" + file)
                    print("deleted file " + dirc + "\\" + file)
        for (dirpath, dirnames, filenames) in os.walk(path_name):
            for dirc in dirnames:
                if dirc[0:8] != 'dir_aip_':
                    continue
                if os.path.isfile(path_name + "\\" + dirc + "\\aip.log"):
                    os.remove(path_name + "\\" + dirc + "\\aip.log")
                    print("deleted file " + dirc + "\\aip.log")
                clear_dir(path_name + "\\" + dirc + "\\play")
    def getShortLog(self, ):
        args = []
        args.append('-getShortLog')
        p = self.createAnswerPipe()
        pipe_descr = self.pipeReq(' '.join(args))
        answer = self.getAnswer(p)
        # print (time.asctime( time.localtime( time.time() ) ) + " : " + answer[0: -1])
        return answer.decode('cp1251')
    def log_print(self, mes):
        if type(mes) != str:
            mes = str(mes)
        print (mes)
        if self.log_file != "":
            try:
                myfile  = open(self.log_file, "a")
            except:
                return mes
            global cur_time
            personal_time = time.strftime("%H:%M:%S - ", time.gmtime(time.time()+10800))
            myfile.write(personal_time + personal_time + mes + "\n")
            myfile.close()
        return mes
if __name__ == "__main__":
    main()
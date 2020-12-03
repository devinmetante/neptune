import pathlib
from datetime import datetime

import win32com.client as client


class COMBridge:
    __program_id = ''
    __cls_id = ''
    obj = None

    def __init__(self, pid=__program_id):
        self.obj = client.Dispatch(pid)


PROGID = 'Neptune.NeptuneCtrl.1'
CLSID = '{50182A02-9665-4094-AA35-B1938D210DAC}'
PARENT_PATH = pathlib.Path(__file__).parent
IMAGE_PATH = PARENT_PATH / 'img'
RES_PATH = PARENT_PATH / 'res'


class Neptune(COMBridge):
    __prog_id = PROGID
    __cls_id = CLSID
    __cam = None
    __cam_ind = None

    def __init__(self):
        super().__init__(pid=self.__prog_id)
        self.__cam = self.obj
        del self.obj

    @staticmethod
    def __prep_paths():
        IMAGE_PATH.mkdir(parents=True, exist_ok=True)
        RES_PATH.mkdir(parents=True, exist_ok=True)

    def initialize_camera(self, cam_ind, pixel_format_ind):
        # cam_ind from camera_list
        # pixel_format_ind from pixel_format_list
        self.camera = cam_ind
        self.pixel_format = pixel_format_ind

        _ = self.camera_type
        if _ == 0:
            self.__cam.GigeFrameRate = 30
        elif _ == 1:
            self.__cam.FireWireFrameRate = 2
        elif _ == 2:
            self.__cam.USBFrameRate = 30

    @property
    def camera_list(self):
        return self.__cam.GetCameraList()

    @property
    def pixel_format_list(self):
        return self.__cam.GetPixelFormatList()

    @property
    def avi_codec_list(self):
        return self.__cam.GetAVICodecList()

    @property
    def camera_info(self):
        # only for GigE camera
        # Select/Set the camera first using obj.camera setter to get the camera info of selected camera
        _ = self.__cam.GetCameraInfo(self.camera)
        return {'model': _[0],
                'vendor': _[1],
                'sn': _[2],
                'user_id': _[3],
                'gateway_address': _[4],
                'ip_address': _[5],
                'mac_address': _[6],
                'subnet_mask': _[7],
                'configuration_mode': _[8],  # 0-Persistent, 1-DHCP
                'nic_ip_address': _[9],
                'nic_subnet_mask': _[10]}

    @property
    def camera_type(self):
        return self.__cam.GetCameraType()

    @property
    def camera(self):
        return self.__cam.Camera

    @camera.setter
    def camera(self, val):
        if isinstance(val, int):
            self.__cam.Camera = val
        elif isinstance(val, str):
            self.__cam.CameraUserID = val
        else:
            raise Exception('Invalid argument type. Expected int or str but received {}.'.format(type(val)))

    @property
    def pixel_format(self):
        return self.__cam.PixelFormat

    @pixel_format.setter
    def pixel_format(self, ind):
        # Disable the camera acquisition before updating the camera pixel format
        # Set camera acquisition to previous setting after updating the camera pixel format
        _ = self.acquisition
        self.acquisition = 0
        self.__cam.PixelFormat = ind
        self.acquisition = _

    @property
    def acquisition(self):
        return self.__cam.Acquisition

    @acquisition.setter
    def acquisition(self, val):
        assert val in (0, 1)
        self.__cam.Acquisition = val

    @property
    def acquisition_mode(self):
        return self.__cam.AcquisitionMode

    @acquisition_mode.setter
    def acquisition_mode(self, str_mode):
        assert str_mode in ('SingleFrame', 'MultiFrame', 'Continuous')
        self.__cam.AcquisitionMode = str_mode

    @property
    def access_mode(self):
        return self.__cam.AccessMode

    @access_mode.setter
    def access_mode(self, val):
        # val: 0-Exclusive, 1-Control, 2-Monitor
        assert val in (0, 1, 2)
        self.__cam.AccessMode = val

        # for GigE Camera only
        if self.camera_type == 1:
            self.__stream_mode = 0 if not val else 1

    @property
    def event_channel(self):
        return self.__cam.EventChannel

    @event_channel.setter
    def event_channel(self, val):
        assert val in (0, 1)
        self.__cam.EventChannel = val

    @property
    def __stream_mode(self):
        return self.__cam.StreamMode

    @__stream_mode.setter
    def __stream_mode(self, val):
        assert val in (0, 1)
        self.__cam.StreamMode = val

    @property
    def __data_bit(self):
        return self.__cam.DataBit

    @__data_bit.setter
    def __data_bit(self, val):
        self.__cam.DataBit = val

    @property
    def __bit_per_pixel(self):
        return self.__cam.GetBitPerPixel()

    @property
    def __sizeX(self):
        return self.__cam.SizeX

    @__sizeX.setter
    def __sizeX(self, val):
        self.__cam.SizeX = val

    @property
    def __sizeY(self):
        return self.__cam.SizeY

    @__sizeY.setter
    def __sizeY(self, val):
        self.__cam.SizeY = val

    @property
    def raw_data(self):
        return self.__cam.GetRawData()

    @property
    def rgb_data(self):
        return self.__cam.GetRGBData()

    @property
    def image_time_stamp(self):
        return self.__cam.GetTimeStamp()

    def save_camera_parameter(self, target_file='Param.txt'):
        # camera parameter file will be saved in RES_PATH directory
        path = RES_PATH / target_file
        self.__cam.SaveCameraParameter(target_file)
        pathlib.Path(target_file).replace(path)

    def load_camera_parameter(self, source_file='Param.txt'):
        # source_file should be initially saved in RES_PATH directory
        path = RES_PATH / source_file
        path.replace(source_file)
        try:
            self.__cam.LoadCameraParameter(source_file)
        finally:
            pathlib.Path(source_file).replace(path)

    @property
    def avi_codec(self):
        return self.__cam.AVICodec

    @avi_codec.setter
    def avi_codec(self, val):
        self.__cam.AVICodec = val

    def save_image(self, img_type):
        assert img_type in ('raw', 'rgb', 'bmp', 'jpg', 'tif')
        # self.acquisition = 1
        self.__cam.Grab()
        datetime_stamp = datetime.now()
        filename = '{}_{:0>2}_{:0>2}_{:0>2}_{:0>2}_{:0>2}_{:0>3}.{}'.format(datetime_stamp.year, datetime_stamp.month,
                                                                            datetime_stamp.day, datetime_stamp.hour,
                                                                            datetime_stamp.minute,
                                                                            datetime_stamp.second,
                                                                            int(datetime_stamp.microsecond / 1000),
                                                                            img_type)
        file_pathname = pathlib.Path(IMAGE_PATH) / filename
        if img_type == 'raw':
            _ = self.__cam.GetRawData(0)
            bytes_per_pixel = self.__bit_per_pixel / 8
            assert file_pathname.write_bytes(_) == (bytes_per_pixel * self.__sizeX * self.__sizeY)
        elif img_type == 'rgb':
            _ = self.__cam.GetRGBData(0)
            assert file_pathname.write_bytes(_) == (3 * self.__sizeX * self.__sizeY)
        else:
            self.__cam.SaveImage(file_pathname, 100)

        return str(file_pathname)

    @property
    def grab_time_out(self):
        return self.__cam.GrabTimeOut

    @grab_time_out.setter
    def grab_time_out(self, val):
        self.__cam.GrabTimeOut = val

    @property
    def error(self):
        return self.__cam.GetError()


if __name__ == '__main__':
    cam = Neptune()
    print('...getting camera list')
    cam_list = cam.camera_list()
    print('\nCAM LIST\n{}'.format(cam_list))
    assert len(cam_list) > 0
    print('...selecting first camera')
    cam.camera = 0
    cam_type = {0: 'GigE', 1: '1394', 2: 'USB3'}
    print('CAMERA TYPE:\t{}'.format(cam_type[cam.camera_type]))
    print('CAMERA INFO:\t{}'.format(cam.camera_info))
    print('Saving Camera Parameter to file...')
    cam.save_camera_parameter()
    print('\nPIXEL FORMAT LIST\n{}'.format(cam.pixel_format_list))
    print('\nPIXEL FORMAT: \t{}'.format(cam.pixel_format))
    print('ACQUISITION:\t{}'.format(cam.acquisition))
    print('ACQUISITION MODE:\t{}'.format(cam.access_mode))
    print('\n...acquiring images...\n')
    print('SAVED\t{}'.format(cam.save_image('jpg')))
    print('SAVED\t{}'.format(cam.save_image('bmp')))
    print('SAVED\t{}'.format(cam.save_image('tif')))
    print('SAVED\t{}'.format(cam.save_image('raw')))
    print('SAVED\t{}'.format(cam.save_image('rgb')))

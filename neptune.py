from harvesters.core import Harvester
import logging
import matplotlib.pyplot as plt

import os
print(os.getenv('HARVESTERS_XML_FILE_DIR'))

# set up a logger for the harvester library.
# this is not needed but can be useful for debugging your script
logger = logging.getLogger('harvesters')
ch = logging.StreamHandler()
logger.setLevel(logging.DEBUG)
ch.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
ch.setFormatter(formatter)
logger.addHandler(ch)

# Create the harvester
h = Harvester(logger=logger)
# the harvester can load dlls as well as cti files.
h.add_cti_file(r'C:\Program Files\IMI Tech\Neptune\ActiveX\Lib\NeptuneTL.cti')
h.update_device_info_list()
print(h.device_info_list)

# create an image acquirer
ia = h.create_image_acquirer(list_index=0)
# this is required for larger images (> 16 MiB) with Critical Link's producer.
ia.num_buffers = 4
ia.remote_device.node_map.PixelFormat.value = 'Mono8'
# Uncomment to set the image ROI width, height. Otherwise will get full frame
# ia.remote_device.node_map.Width.value, ia.remote_device.node_map.Height.value = 800, 600

print("Starting Acquistion")

ia.start_image_acquisition()

# just capture 1 frame
for i in range(1):
	with ia.fetch_buffer(timeout=4) as buffer:
		payload = buffer.payload
		component = payload.components[0]
		width = component.width
		height = component.height
		data_format = component.data_format
		print("Image details: {}w {}h {}".format(width, height, data_format))
		# for monochrome 8 bit images
		if int(component.num_components_per_pixel) == 1:
			content = component.data.reshape(height, width)
		else:
			content = component.data.reshape(height, width, int(component.num_components_per_pixel))
		if int(component.num_components_per_pixel) == 1:
			plt.imshow(content, cmap='gray')
		else:
			plt.imshow(content)
		plt.show()

#
ia.stop_image_acquisition()
ia.destroy()

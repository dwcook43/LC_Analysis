import setuptools

setuptools.setup(
    #Needed to silence warnings (and to be a worthwhile package)
    name='LC_analysis',
    url='https://github.com/dwcook43/LC_Analysis',
    author='Daniel W. Cook & Ryan Nelson',
    author_email='dwcookphd@gmail.com',
    # Needed to actually package something
    packages=['LC_Report']
    # Needed for dependencies
    install_requires = [
        'numpy>=1.17',
        'pandas>=0.25' ],    # *strongly* suggested for sharing
    version='0.1',
    # The license can be anything you like
    license='GNUv3',
    description='Tool for visualize and reporting Agilent HPLC data',
    # We will also need a readme eventually (there will be a warning)
    # long_description=open('README.txt').read(),
)

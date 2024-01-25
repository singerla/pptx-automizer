import * as ModifyBackground from '../src/modify/modify-background'; // Adjust
import Automizer from '../src/automizer';

describe('ModifyMasterBackground', () => {
    beforeEach(() => {
        // Initialize the automizer instance
        const automizer = new Automizer({
            templateDir: `${__dirname}/pptx-templates`,
            outputDir: `${__dirname}/pptx-output`,
          });
        
        const pres = await automizer
        .loadRoot(`EmptyTemplate.pptx`)
    });

    test('should allow setting a background color on a master slide', async () => {
        const masterId = 1; // Example master slide ID
        const newColor = 'FF5733'; // Example new color

        // Modify the master slide by setting a new background color
        await pres.addMaster(masterId, (master) => {
            master.modifyBackground.setBackgroundColor(master, newColor);
        });

        // Later on we could add a way to getMasterInfo or something that could either be the raw xml, or more helpfully a {background:{backgroundImage:"",backgroundColor:""}}
        // Retrieve the modified master slide and verify the background color
        const modifiedMaster = await pres.getMasterInfo(masterId);
        expect(modifiedMaster.backgroundColor).toBe(newColor); // Assuming the master slide object has a backgroundColor property
    });

    test('should allow setting a background image on a master slide', async () => {
        const masterId = 1; // Example master slide ID
        const newImage = 'path/to/new/image.jpg'; // Example new image path

        // Modify the master slide by setting a new background image
        await automizer.addMaster(masterId, (master) => {
            master.modifyBackground.setBackgroundImage(master, newImage);
        });
        // Later on we could add a way to getMasterInfo or something that could either be the raw xml, or more helpfully a {background:{backgroundImage:"",backgroundColor:""}}
        // Retrieve the modified master slide and verify the background image
        const modifiedMaster = await pres.getMasterInfo(masterId);
        expect(modifiedMaster.backgroundImage).toBe(newImage); // Assuming the master slide object has a backgroundImage property
    });

    // Additional tests for edge cases, error handling, etc.
});

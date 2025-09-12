import Automizer, { modify, read, XmlHelper } from '../src/index';
import { vd } from '../src/helper/general-helper';
import { XmlSlideHelper } from '../src/helper/xml-slide-helper';
import { ElementInfo } from '../src/types/xml-types';

test('read shape group info', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    verbosity: 2,
    removeExistingSlides: true,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`SlideWithShapes.pptx`);

  pres.addSlide("SlideWithShapes.pptx", 3, async (slide) => {
    // Get info for top level group
    const infoTopLevel = await slide.getElement('TopLevelGroup');
    const groupInfoTopLevel = infoTopLevel.getGroupInfo();

    // Get info for sub group
    const infoSubGroup = await slide.getElement('Subgroup');
    const groupInfoSubGroup = infoSubGroup.getGroupInfo();

    // Get info for shapes in top level group
    const infoDrum = await slide.getElement('Drum');
    const groupInfoDrum = infoDrum.getGroupInfo();

    const infoCloud = await slide.getElement('Cloud');
    const groupInfoCloud = infoCloud.getGroupInfo();

    // Get info for shapes in sub group
    const infoStar = await slide.getElement('Star');
    const groupInfoStar = infoStar.getGroupInfo();

    const infoArrow = await slide.getElement('Arrow');
    const groupInfoArrow = infoArrow.getGroupInfo();

    // Test parent group properties
    expect(groupInfoTopLevel.isParent).toBe(true);
    expect(groupInfoTopLevel.isChild).toBe(false);
    expect(groupInfoTopLevel.getChildren().length).toBeGreaterThan(0);

    // Test nested group properties
    expect(groupInfoSubGroup.isParent).toBe(true);
    expect(groupInfoSubGroup.isChild).toBe(true);

    // Test that SubGroup's parent is TopLevelGroup
    const subGroupParent = groupInfoSubGroup.getParent();
    expect(subGroupParent).not.toBeNull();
    expect(XmlSlideHelper.getElementName(subGroupParent)).toBe('TopLevelGroup');

    // Test that Drum is a child of TopLevelGroup
    expect(groupInfoDrum.isChild).toBe(true);
    expect(groupInfoDrum.isParent).toBe(false);
    const drumParent = groupInfoDrum.getParent();
    expect(drumParent).not.toBeNull();
    expect(XmlSlideHelper.getElementName(drumParent)).toBe('TopLevelGroup');

    // Test that Cloud is a child of TopLevelGroup
    expect(groupInfoCloud.isChild).toBe(true);
    expect(groupInfoCloud.isParent).toBe(false);
    const cloudParent = groupInfoCloud.getParent();
    expect(cloudParent).not.toBeNull();
    expect(XmlSlideHelper.getElementName(cloudParent)).toBe('TopLevelGroup');

    // Test that Star is a child of SubGroup
    expect(groupInfoStar.isChild).toBe(true);
    expect(groupInfoStar.isParent).toBe(false);
    const starParent = groupInfoStar.getParent();
    expect(starParent).not.toBeNull();
    expect(XmlSlideHelper.getElementName(starParent)).toBe('Subgroup');

    // Test that Arrow is a child of SubGroup
    expect(groupInfoArrow.isChild).toBe(true);
    expect(groupInfoArrow.isParent).toBe(false);
    const arrowParent = groupInfoArrow.getParent();
    expect(arrowParent).not.toBeNull();
    expect(XmlSlideHelper.getElementName(arrowParent)).toBe('Subgroup');

    // Test that TopLevelGroup's children include SubGroup, Drum, and Cloud
    const topLevelChildren = groupInfoTopLevel.getChildren();
    const topLevelChildrenNames = topLevelChildren.map(child =>
      XmlSlideHelper.getElementName(child)
    );
    expect(topLevelChildrenNames).toContain('Subgroup');
    expect(topLevelChildrenNames).toContain('Drum');
    expect(topLevelChildrenNames).toContain('Cloud');

    // Test that SubGroup's children include Star and Arrow
    const subGroupChildren = groupInfoSubGroup.getChildren();
    const subGroupChildrenNames = subGroupChildren.map(child =>
      XmlSlideHelper.getElementName(child)
    );
    expect(subGroupChildrenNames).toContain('Star');
    expect(subGroupChildrenNames).toContain('Arrow');
  });

  await pres.write(`read-group-info.test.pptx`);
});


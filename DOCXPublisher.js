const {
  Packer,
  Header,
  Paragraph,
  ImageRun,
  TextRun,
  File,
  AlignmentType,
  StyleLevel,
  TableOfContents,
  HeadingLevel,
  HorizontalPositionRelativeFrom,
  VerticalPositionRelativeFrom,
  HorizontalPositionAlign,
  VerticalPositionAlign,
  TextWrappingType,
  TextWrappingSide,
  FrameAnchorType,
  FrameWrap,
  Numbering,
  Document,
  LevelFormat,
  LineNumberRestartFormat,
  Table,
  TableRow,
  TableCell,
  VerticalAlign,
  TextDirection,
  convertInchesToTwip,
  convertMillimetersToTwip,
  BorderStyle,
  WidthType,
  Footer,
  PageNumber,
  TabStopType,
  TabStopPosition,
  UnderlineType,
  PageNumberFormat,
} = require("docx");
const jsdom = require("jsdom");
const { JSDOM } = jsdom;
const async = require("async");
const path = require("path");
const fs = require("fs-extra");
const { htmlToText } = require("html-to-text");
const cheerio = require("cheerio");
const usermanager = require("../../usermanager");
const configuration = require("../../configuration");
const { Constants } = require("../../outputmanager");
const PublishDataQueries = require("../query");
const CourseTransformer = require("../transformers/CourseTransformer");
const logger = require("../../logger");

class DOCXPublisher {
  constructor() {
    this.publishQueries = new PublishDataQueries();
    this.courseTransformer = new CourseTransformer();
    this.ckeditorBuilder = new CKEditorTransformer();
    this.creationDate = new Date();
    this.configs = {
      issueCode: "PJT-325454",
      issueVersion: "v3.0",
      contentGroup: "Content Group (N/A)",
      disclaimer: "FOR TRAINING USE ONLY",
      disclosure:
        "Use and disclosure of this document is controlled; see Title/Cover page.",
      contentType: "Student Notes",
      chapters: "(ATA Chapters: 29, 52)",
      copyright: "© Copyright 2021 Leonardo MW Ltd",
      notice:
        "This document contains information that is confidential and proprietary to Leonardo MW Ltd (‘the Company’) and is supplied on the express condition that it may not be disclosed to any third party, or reproduced in whole or in part, or used for manufacture, or used for any purpose other than for which it is supplied, without the prior written consent of the Company. Every permitted disclosure, reproduction, adaptation or publication of this document in whole or in part and in any manner or form shall prominently contain this notice.",
      font: "Arial",
      divider:
        "_________________________________________________________________________",
    };
    this.assets = [];
  }

  async build(courseId, contentGroup) {
    try {
      const manualDirPath = this.getTargetDir(courseId);
      await this.prepareBuildDir(manualDirPath);
      const course = await this.publishQueries.getCourse(
        courseId,
        contentGroup
      );
      const themeAttr = course.theme.targetAttribute;
      this.themeSettings =
        course.themeSettings && course.themeSettings[themeAttr];

      this.assets = course.assets.records || [];
      const manualData = await this.courseTransformer.getCourseTreeJSON(course);
      const fileName = await this.generateDocX(
        manualData,
        manualDirPath,
        course.assets.records,
        contentGroup
      );
      return fileName;
    } catch (error) {
      console.log(error);
      throw error;
    }
  }

  getTargetDir(courseId) {
    const user = usermanager.getCurrentUser();
    const tenantId = user.tenant._id;
    const dirPath = path.join(
      configuration.tempDir,
      configuration.getConfig("masterTenantID"),
      Constants.Folders.Framework,
      Constants.Folders.AllCourses,
      tenantId,
      courseId,
      Constants.Folders.DOCX
    );
    return dirPath;
  }

  getAssetDir(targetDir) {
    return path.join(
      targetDir,
      Constants.Folders.Course,
      Constants.Folders.Assets
    );
  }

  getDefaultAssetOrFromTheme() {
    return "path";
  }

  getLogoPaths() {
    if (this.themeSettings && this.themeSettings._supportForManual) {
    }
  }

  resolveThemeManualAssets() {
    const themePath = path.join(configuration.getConfig("root"), "");
  }

  loadManualLogs() {
    const logo1 = new ImageRun({
      data: fs.readFileSync(path.join(__dirname, "assets", "logo_1.png")),
      transformation: {
        height: 130.39370079,
        width: 642.14173228,
      },
      floating: {
        zIndex: 1,
        horizontalPosition: {
          // relative: HorizontalPositionRelativeFrom.COLUMN,
          align: HorizontalPositionAlign.CENTER,
        },
        verticalPosition: {
          relative: VerticalPositionRelativeFrom.PARAGRAPH,
          align: VerticalPositionAlign.TOP,
        },
        wrap: {
          type: TextWrappingType.SQUARE,
          side: TextWrappingSide.BOTH_SIDES,
        },
      },
    });

    const logo2 = new ImageRun({
      data: fs.readFileSync(path.join(__dirname, "assets", "logo_2.png")),
      transformation: {
        height: 205.22834646,
        width: 641.00787402,
      },
      floating: {
        zIndex: 1,
        horizontalPosition: {
          // relative: HorizontalPositionRelativeFrom.COLUMN,
          align: HorizontalPositionAlign.CENTER,
        },
        verticalPosition: {
          relative: VerticalPositionRelativeFrom.PARAGRAPH,
          align: VerticalPositionAlign.TOP,
        },
        wrap: {
          type: TextWrappingType.SQUARE,
          side: TextWrappingSide.BOTH_SIDES,
        },
      },
    });

    const logo3 = new ImageRun({
      data: fs.readFileSync(path.join(__dirname, "assets", "logo_3.png")),
      transformation: {
        height: 65.385826772,
        width: 679.93700787,
      },
      floating: {
        zIndex: 1,
        horizontalPosition: {
          // relative: HorizontalPositionRelativeFrom.COLUMN,
          align: HorizontalPositionAlign.CENTER,
        },
        verticalPosition: {
          relative: VerticalPositionRelativeFrom.PARAGRAPH,
          align: VerticalPositionAlign.TOP,
        },
        wrap: {
          type: TextWrappingType.SQUARE,
          side: TextWrappingSide.BOTH_SIDES,
        },
      },
    });

    const logo4 = new ImageRun({
      data: fs.readFileSync(path.join(__dirname, "assets", "logo_4.png")),
      transformation: {
        height: 87.307086614,
        width: 640.31496063,
      },
      floating: {
        zIndex: 1,
        horizontalPosition: {
          align: HorizontalPositionAlign.CENTER,
        },
        verticalPosition: {
          offset: 9470000,
        },
      },
    });

    const logo5 = new ImageRun({
      data: fs.readFileSync(path.join(__dirname, "assets", "logo_5.png")),
      transformation: {
        height: 63.118110236,
        width: 314.07874016,
      },
      floating: {
        horizontalPosition: {
          align: HorizontalPositionAlign.LEFT,
          relative: HorizontalPositionRelativeFrom.MARGIN,
        },
        verticalPosition: {
          relative: VerticalPositionRelativeFrom.PARAGRAPH,
          align: VerticalPositionAlign.TOP,
        },
        wrap: {
          type: TextWrappingType.SQUARE,
          side: TextWrappingSide.BOTH_SIDES,
        },
      },
    });
    return [logo1, logo2, logo3, logo4, logo5];
  }

  getBorderOptions() {
    const border = {
      color: "auto",
      space: 1,
      value: "single",
      size: 6,
    };
    return {
      top: border,
      bottom: border,
      left: border,
      right: border,
    };
  }

  async prepareBuildDir(dirPath) {
    try {
      await fs.ensureDir(dirPath);
      await fs.emptyDir(dirPath);
    } catch (error) {
      logger.log("error", error);
      throw error;
    }
  }

  findAssetById(assets, assetSource) {
    const assetFileName = assetSource.replace("course/assets/", "");
    const asset = assets.find((asset) => {
      return assetFileName === asset.filename;
    });
    return asset;
  }

  buildAssetPath(asset) {
    const user = usermanager.getCurrentUser();
    const tenantName = user.tenant.name;
    return path.join(
      configuration.getConfig("root"),
      configuration.getConfig("dataRoot"),
      tenantName,
      asset.path
    );
  }

  generateDisclaimer() {
    return new TextRun({
      text: this.configs.disclaimer,
      bold: true,
      font: this.configs.font,
      size: "11pt",
    });
  }

  generateBlankPage() {
    return new Paragraph({
      pageBreakBefore: true,
      alignment: AlignmentType.CENTER,
      frame: {
        wrap: FrameWrap.NONE,
        position: {
          x: 0,
          y: 7029.9212598,
        },
        width: 0,
        height: 0,
        anchor: {
          horizontal: FrameAnchorType.MARGIN,
          vertical: FrameAnchorType.PAGE,
        },
        alignment: {
          x: HorizontalPositionAlign.CENTER,
        },
      },
      children: [
        new TextRun({
          text: "INTENTIONALLY",
          bold: false,
          font: this.configs.font,
          size: "24pt",
          color: "#C0C0C0",
        }),
        new TextRun({
          text: "LEFT",
          bold: false,
          font: this.configs.font,
          size: "24pt",
          break: true,
          color: "#C0C0C0",
        }),
        new TextRun({
          text: "BLANK",
          bold: false,
          font: this.configs.font,
          size: "24pt",
          break: true,
          color: "#C0C0C0",
        }),
      ],
    });
  }

  generateLogoFour(logo4) {
    return new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [logo4],
    });
  }

  async generateDocX(course, targetDir, assets, contentGroup) {
    const children = [];
    const [logo1, logo2, logo3, logo4, logo5] = this.loadManualLogs();
    const border = this.getBorderOptions();
    this.iterateContentObjects(course, children);
    const doc = new Document({
      features: {
        updateFields: true,
      },
      numbering: {
        config: [
          {
            reference: "content",
            levels: [
              {
                level: 0,
                format: LevelFormat.DECIMAL,
                text: "%1",
              },
              {
                level: 1,
                format: LevelFormat.DECIMAL,
                text: "%1.%2",
              },
              {
                level: 2,
                format: LevelFormat.DECIMAL,
                text: "%1.%2.%3",
              },
            ],
          },
          {
            reference: "bullet-points",
            levels: [
              {
                level: 0,
                format: LevelFormat.BULLET,
                text: "\u2022",
                alignment: AlignmentType.LEFT,
                style: {
                  paragraph: {
                    indent: {
                      left: convertInchesToTwip(0.5),
                      hanging: convertInchesToTwip(0.25),
                    },
                  },
                },
              },
              {
                level: 1,
                format: LevelFormat.BULLET,
                text: "\u2022",
                alignment: AlignmentType.LEFT,
                style: {
                  paragraph: {
                    indent: {
                      left: convertInchesToTwip(1),
                      hanging: convertInchesToTwip(0.25),
                    },
                  },
                },
              },
              {
                level: 2,
                format: LevelFormat.BULLET,
                text: "\u2022",
                alignment: AlignmentType.LEFT,
                style: {
                  paragraph: {
                    indent: { left: 2160, hanging: convertInchesToTwip(0.25) },
                  },
                },
              },
              {
                level: 3,
                format: LevelFormat.BULLET,
                text: "\u2022",
                alignment: AlignmentType.LEFT,
                style: {
                  paragraph: {
                    indent: { left: 2880, hanging: convertInchesToTwip(0.25) },
                  },
                },
              },
              {
                level: 4,
                format: LevelFormat.BULLET,
                text: "\u2022",
                alignment: AlignmentType.LEFT,
                style: {
                  paragraph: {
                    indent: { left: 3600, hanging: convertInchesToTwip(0.25) },
                  },
                },
              },
            ],
          },
        ],
      },
      styles: {
        default: {
          document: {
            run: {
              font: this.configs.font,
            },
          },
        },
        // heading1: {
        //   run: {
        //     size: 28,
        //     bold: true,
        //     italics: true,
        //     color: "FF0000",
        //   },
        //   paragraph: {
        //     spacing: {
        //       after: 120,
        //     },
        //   },
        // },
        // heading2: {
        //   run: {
        //     size: 26,
        //     bold: true,
        //     underline: {
        //       type: UnderlineType.DOUBLE,
        //       color: "FF0000",
        //     },
        //   },
        //   paragraph: {
        //     spacing: {
        //       before: 240,
        //       after: 120,
        //     },
        //   },
        // },
        paragraphStyles: [
          {
            id: "WorkbitHeading1",
            name: "WorkbitHeading1",
            basedOn: "Heading1",
            next: "Heading1",
            quickFormat: true,
            run: {
              font: this.configs.font,
              size: 28,
              bold: true,
              color: "#000000",
            },
          },
          {
            id: "WorkbitHeading2",
            name: "WorkbitHeading2",
            basedOn: "Heading2",
            next: "Heading2",
            quickFormat: true,
            run: {
              font: this.configs.font,
              size: 24,
              bold: true,
              color: "#000000",
            },
          },
          {
            id: "WorkbitHeading3",
            name: "WorkbitHeading3",
            basedOn: "Heading3",
            next: "Heading3",
            quickFormat: true,
            run: {
              font: this.configs.font,
              size: 20,
              bold: true,
              color: "#000000",
            },
          },
        ],
      },
      sections: [
        this.createTitlePage(course),
        this.createIssuePage(course),
        {
          properties: {
            page: {
              pageNumbers: {
                start: 1,
                formatType: PageNumberFormat.DECIMAL,
              },
              margin: {
                top: convertMillimetersToTwip(20),
                bottom: convertMillimetersToTwip(15),
                left: convertMillimetersToTwip(15),
                right: convertMillimetersToTwip(15),
                mirror: true,
              },
            },
          },
          headers: {
            default: this.generateContentHeader(course, contentGroup, logo5),
          },
          footers: this.generateNumberedFooter(),
          children: children,
        },
      ],
    });

    try {
      const buffer = await Packer.toBuffer(doc);
      const fileName = course.title + ".docx";
      const docxPath = path.join(targetDir, fileName);
      await fs.writeFile(docxPath, buffer);
      return fileName;
    } catch (error) {
      throw error;
    }
  }

  iterateContentObjects(course, children) {
    for (const [pageIndex, page] of course.contentObjects.entries()) {
      if (page._type !== "page") {
        continue;
      }
      children.push(
        new Paragraph({
          text: page.displayTitle,
          heading: HeadingLevel.HEADING_1,
          spacing: {
            before: 200,
          },
          style: "WorkbitHeading1",
          pageBreakBefore: true,
          numbering: {
            level: 0,
            reference: "content",
            instance: 1,
          },
        })
      );
      this.iterateArticles(page.articles.entries(), children);
    }
  }

  iterateArticles(articleEntries, children) {
    for (const [articleIndex, article] of articleEntries) {
      children.push(
        new Paragraph({
          text: article.displayTitle,
          style: "WorkbitHeading2",
          spacing: {
            before: 200,
          },
          heading: HeadingLevel.HEADING_2,
          pageBreakBefore: articleIndex !== 0,
          numbering: {
            level: 1,
            reference: "content",
            instance: 1,
          },
        })
      );
      this.iterateBlocks(article.blocks.entries(), children);
    }
  }

  iterateBlocks(blockEntries, children) {
    for (const [blockIndex, block] of blockEntries) {
      children.push(
        new Paragraph({
          text: block.displayTitle,
          heading: HeadingLevel.HEADING_3,
          spacing: {
            before: 200,
          },
          style: "WorkbitHeading3",
          pageBreakBefore: blockIndex !== 0,
          numbering: {
            level: 2,
            reference: "content",
            instance: 1,
          },
        })
      );

      const component = block.components[0];
      this.handleComponent(component, children);
    }
  }

  handleComponent(component, children) {
    switch (component._component) {
      case "graphic": {
        const graphicURI = component.properties.small;
        const asset = this.findAssetById(this.assets, graphicURI);
        if (asset) {
          const assetPath = this.buildAssetPath(asset);
          const height = asset.metadata.height;
          const width = asset.metadata.width;
          const assetImageRun = new ImageRun({
            data: fs.readFileSync(assetPath),
            transformation: {
              height: 500,
              width: 700,
            },
            floating: {
              zIndex: 1,
              horizontalPosition: {
                // relative: HorizontalPositionRelativeFrom.COLUMN,
                align: HorizontalPositionAlign.CENTER,
              },
              verticalPosition: {
                relative: VerticalPositionRelativeFrom.PARAGRAPH,
                align: VerticalPositionAlign.TOP,
              },
              wrap: {
                type: TextWrappingType.SQUARE,
                side: TextWrappingSide.BOTH_SIDES,
              },
            },
          });
          children.push(
            new Paragraph({
              text: "",
              font: this.configs.font,
              size: "11pt",
              break: true,
              children: [assetImageRun],
            })
          );
        }
      }
      default: {
        children.push(...this.translateCKEditorToDocx(component.body));
      }
    }
  }

  translateCKEditorToDocx(body) {
    return this.ckeditorBuilder.transform(body) || [];
  }

  createTitlePage(course) {
    const border = this.getBorderOptions();
    const [logo1, logo2, logo3, logo4, logo5] = this.loadManualLogs();
    return {
      properties: {
        page: {
          margin: {
            top: 1133.8582677,
            bottom: 850.39370079,
            header: 396.8503937,
            footer: 396.8503937,
            gutter: 0,
            mirror: true,
          },
        },
      },
      headers: {
        default: this.generateDefaultHeader(),
      },
      footers: this.generateDefaultFooter(),
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [logo1],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [logo2],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [logo3],
        }),
        new Paragraph({
          frame: {
            position: {
              x: 283.46456693,
              y: 6366.6141732,
            },
            width: 11338.582677,
            height: 3617.007874,
            anchor: {
              horizontal: FrameAnchorType.PAGE,
              vertical: FrameAnchorType.MARGIN,
            },
            alignment: {
              x: HorizontalPositionAlign.CENTER,
              y: VerticalPositionAlign.TOP,
            },
          },
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text: "AW101 NAWSARH",
              bold: true,
              font: this.configs.font,
              size: "26pt",
            }),
            new TextRun({
              text: this.formatContentGroup(course.contentGroup),
              bold: true,
              font: this.configs.font,
              size: "26pt",
              break: true,
            }),
            new TextRun({
              font: this.configs.font,
              size: "26pt",
              text: "",
              break: true,
            }),
            new TextRun({
              text: course.displayTitle,
              bold: true,
              font: this.configs.font,
              size: "26pt",
              break: true,
            }),
            new TextRun({
              text: this.configs.contentType,
              bold: true,
              font: this.configs.font,
              size: "26pt",
              break: true,
            }),
            new TextRun({
              text: this.configs.chapters,
              bold: true,
              font: this.configs.font,
              size: "18pt",
              break: true,
            }),
          ],
        }),

        new Paragraph({
          frame: {
            position: {
              x: 0,
              y: 11366.6141732,
            },
            width: 9500,
            height: 1000.007874,
            anchor: {
              horizontal: FrameAnchorType.MARGIN,
              vertical: FrameAnchorType.PAGE,
            },
            alignment: {
              x: HorizontalPositionAlign.CENTER,
              y: VerticalPositionAlign.CENTER,
            },
          },
          border: border,
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text: this.configs.copyright,
              bold: true,
              font: this.configs.font,
              size: "8pt",
            }),
            new TextRun({
              text: this.configs.notice,
              bold: false,
              font: this.configs.font,
              size: "8pt",
              break: true,
            }),
          ],
        }),
        new Paragraph({
          frame: {
            position: {
              x: 0,
              y: 12600,
            },
            width: 9500,
            height: 1250,
            anchor: {
              horizontal: FrameAnchorType.MARGIN,
              vertical: FrameAnchorType.PAGE,
            },
            alignment: {
              x: HorizontalPositionAlign.CENTER,
              y: VerticalPositionAlign.CENTER,
            },
          },
          border: border,
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text: "",
              bold: false,
              font: this.configs.font,
              size: "8pt",
              break: true,
            }),
          ],
        }),
        new Paragraph({
          frame: {
            position: {
              x: 0,
              y: 13900,
            },
            width: 9500,
            height: 1000,
            anchor: {
              horizontal: FrameAnchorType.MARGIN,
              vertical: FrameAnchorType.PAGE,
            },
            alignment: {
              x: HorizontalPositionAlign.CENTER,
              y: VerticalPositionAlign.CENTER,
            },
          },
          border: border,
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text: "DOCUMENT NO: PJT-325454",
              bold: true,
              font: this.configs.font,
              size: "10pt",
            }),
            new TextRun({
              text: "ISSUE: V3.0",
              bold: true,
              font: this.configs.font,
              size: "10pt",
              break: true,
            }),
            new TextRun({
              text: "ISSUE DATE: 01st October 2019 {T}",
              bold: true,
              font: this.configs.font,
              size: "10pt",
              break: true,
            }),
          ],
        }),
        this.generateLogoFour(logo4),
        this.generateBlankPage(),
      ],
    };
  }

  createIssuePage(course) {
    const border1 = {
      top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    };
    const border2 = {
      top: { style: BorderStyle.SINGLE, size: 1, color: "111111" },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: "111111" },
      left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    };
    return {
      properties: {
        page: {
          pageNumbers: {
            start: 1,
            formatType: PageNumberFormat.LOWER_ROMAN,
          },

          margin: {
            // top: 1133.8582677,
            // bottom: 850.39370079,
            // header: 396.8503937,
            // footer: 396.8503937,
            // gutter: 0,
            // mirror: true,
          },
        },
      },
      headers: {
        default: this.generateDefaultHeader(),
      },
      footers: this.generateNumberedFooter(),
      children: [
        new Paragraph({
          alignment: AlignmentType.LEFT,
          children: [
            new TextRun({
              text: "Issue Record",
              bold: true,
              font: this.configs.font,
              size: "14pt",
            }),
            new TextRun({
              text: "The latest issue number shown below wholly defines the standard of this document. Changes to this document will be made by re-issue of the document in its entirety or by amendment list action that will raise the issue.",
              bold: false,
              font: this.configs.font,
              size: "10pt",
              break: true,
            }),
            new TextRun({
              text: "",
              bold: false,
              font: this.configs.font,
              size: "10pt",
              break: true,
            }),
          ],
        }),
        new Table({
          columnWidths: [3505, 5505],
          alignment: AlignmentType.CENTER,
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  borders: border2,
                  width: {
                    size: 5505,
                    type: WidthType.DXA,
                  },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "ISSUE",
                          bold: true,
                          size: "12pt",
                          font: this.configs.font,
                        }),
                      ],
                    }),
                  ],
                }),
                new TableCell({
                  borders: border2,
                  width: {
                    size: 5505,
                    type: WidthType.DXA,
                  },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "DATE",
                          bold: true,
                          size: "12pt",
                          font: this.configs.font,
                        }),
                      ],
                    }),
                  ],
                }),
                new TableCell({
                  borders: border2,
                  width: {
                    size: 5505,
                    type: WidthType.DXA,
                  },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "REMARKS",
                          bold: true,
                          size: "12pt",
                          font: this.configs.font,
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  borders: border1,
                  children: [
                    new Paragraph({
                      text: "1.0",
                    }),
                  ],
                }),
                new TableCell({
                  borders: border1,
                  children: [new Paragraph({ text: "23/02/2017" })],
                }),
                new TableCell({
                  borders: border1,
                  children: [new Paragraph({ text: "" })],
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  borders: border1,
                  children: [
                    new Paragraph({
                      text: "1.0",
                    }),
                  ],
                }),
                new TableCell({
                  borders: border1,
                  children: [new Paragraph({ text: "23/02/2017" })],
                }),
                new TableCell({
                  borders: border1,
                  children: [new Paragraph({ text: "" })],
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  borders: border1,
                  children: [
                    new Paragraph({
                      text: "2.0",
                    }),
                  ],
                }),
                new TableCell({
                  borders: border1,
                  children: [new Paragraph({ text: "23/02/2018" })],
                }),
                new TableCell({
                  borders: border1,
                  children: [new Paragraph({ text: "" })],
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  borders: border1,
                  children: [
                    new Paragraph({
                      text: "3.0",
                    }),
                  ],
                }),
                new TableCell({
                  borders: border1,
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "23/07/2021",
                          font: this.configs.font,
                        }),
                      ],
                    }),
                  ],
                }),
                new TableCell({
                  borders: border1,
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Update to Base 2 Standard",
                          font: this.configs.font,
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
        this.generateBlankPage(),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Table of Contents", size: 28, bold: true }),
          ],
        }),
        new TableOfContents("Table of Contents", {
          hyperlink: true,
          headingStyleRange: "1-3",
          captionLabelIncludingNumbers: true,
          stylesWithLevels: [
            new StyleLevel("WorkbitHeading1", 1),
            new StyleLevel("WorkbitHeading2", 2),
            new StyleLevel("WorkbitHeading3", 3),
          ],
        }),
      ],
    };
  }

  generateDefaultHeader() {
    return new Header({
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text: "",
              bold: true,
              font: this.configs.font,
              size: "11pt",
              break: true,
            }),
            this.generateDisclaimer(),
          ],
        }),
      ],
    });
  }

  generateContentHeader(course, contentGroup, logo5) {
    return new Header({
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text: this.configs.disclaimer,
              bold: true,
              font: this.configs.font,
              size: "11pt",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.LEFT,
          children: [logo5],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [
            new TextRun({
              text: "AW101 NAWSARH",
              bold: true,
              font: this.configs.font,
              size: "11pt",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [
            new TextRun({
              text: course.displayTitle,
              bold: true,
              font: this.configs.font,
              size: "11pt",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [
            new TextRun({
              text: `Student Notes - ${this.formatContentGroup(contentGroup)}`,
              bold: true,
              font: this.configs.font,
              size: "11pt",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text: this.configs.divider,
              bold: true,
              font: this.configs.font,
              size: "11pt",
            }),
            new TextRun({
              text: "",
              bold: true,
              font: this.configs.font,
              size: "10pt",
              break: true,
            }),
          ],
        }),
      ],
    });
  }

  generateDefaultFooter() {
    return {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [this.generateDisclaimer()],
          }),
        ],
      }),
    };
  }

  generateNumberedFooter() {
    const dividerLine = new ImageRun({
      data: fs.readFileSync(path.join(__dirname, "assets", "divider.png")),
      transformation: {
        height: 1,
        width: 700,
      },
      floating: {
        zIndex: 1,
        horizontalPosition: {
          // relative: HorizontalPositionRelativeFrom.COLUMN,
          align: HorizontalPositionAlign.CENTER,
        },
        verticalPosition: {
          relative: VerticalPositionRelativeFrom.PARAGRAPH,
          align: VerticalPositionAlign.TOP,
        },
        wrap: {
          type: TextWrappingType.SQUARE,
          side: TextWrappingSide.BOTH_SIDES,
        },
      },
    });

    return {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: this.configs.divider })],
          }),
          new Paragraph({
            tabStops: [
              {
                type: TabStopType.RIGHT,
                position: TabStopPosition.MAX,
              },
            ],
            alignment: AlignmentType.LEFT,
            children: [
              new TextRun({
                tabStops: [
                  {
                    type: TabStopType.RIGHT,
                    position: TabStopPosition.MAX,
                  },
                ],
                text: `${this.configs.issueCode}`,
                bold: true,
                font: this.configs.font,
                size: "11pt",
              }),
              new TextRun({
                text:
                  "\t\t" +
                  `${this.creationDate.getUTCDate()}/${this.creationDate.getUTCMonth()}/${this.creationDate.getUTCFullYear()}`,
                bold: true,
                font: this.configs.font,
                size: "11pt",
              }),
            ],
          }),
          new Paragraph({
            tabStops: [
              {
                type: TabStopType.RIGHT,
                position: TabStopPosition.MAX,
              },
            ],
            alignment: AlignmentType.LEFT,
            children: [
              new TextRun({
                text: "Issue " + this.configs.issueVersion,
                bold: true,
                font: this.configs.font,
                size: "11pt",
              }),
              new TextRun({
                children: ["\t\tPage ", PageNumber.CURRENT],
                bold: true,
                font: this.configs.font,
                size: "11pt",
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: this.configs.disclosure,
                bold: false,
                color: "#969696",
                font: this.configs.font,
                size: "8pt",
                break: true,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: this.configs.disclaimer,
                bold: true,
                font: this.configs.font,
                size: "11pt",
              }),
            ],
          }),
        ],
      }),
    };
  }

  formatContentGroup(contentGroup) {
    if (!contentGroup || contentGroup === "undefined") {
      return this.configs.contentGroup;
    }
    return contentGroup[0].toUpperCase() + contentGroup.substr(1);
  }
}

class CKEditorTransformer {
  constructor() {
    this.mappableElements = {
      UL: "ul",
      OL: "ol",
      LI: "li",
      P: "p",
      SPAN: "span",
      BR: "br",
      TABLE: "table",
      TEXT: "#text",
    };
    this.applicableStyles = {
      MARGIN: "margin",
      MARGIN_TOP: "margin-top",
      MARGIN_BOTTOM: "margin-bottom",
      MARGIN_LEFT: "margin-left",
      MARGIN_RIGHT: "margin-right",
      FONT_FAMILY: "font-family",
      FONT_SIZE: "font-size",
    };
  }

  transform(html) {
    const docxObjects = [];
    const dom = new JSDOM(`<body>${html}</body>`);
    const Window = dom.window;
    const Doc = Window.document;
    const NodeFilter = dom.window.NodeFilter;
    const firstLevelNodes = Array.from(Doc.body.children);

    for (const firstLevelNode of firstLevelNodes) {
      const nodeName = firstLevelNode.nodeName.toLowerCase();
      //  console.log(`START------------->${nodeName}`);
      switch (nodeName) {
        case this.mappableElements.P: {
          //const p = new Paragraph({});
          const iterator = Doc.createNodeIterator(
            firstLevelNode,
            NodeFilter.SHOW_ALL
          );
          let currentNode = iterator.nextNode();
          let temp = [];
          let stack = [];
          while ((currentNode = iterator.nextNode())) {
            const currentNodeName = currentNode.nodeName.toLowerCase();
            if (currentNodeName == this.mappableElements.TEXT) {
              const parent = currentNode.parentElement;
              let style1 = Window.getComputedStyle(parent);
              const color = style1.getPropertyValue("color");
              const foregroundColor = style1.getPropertyValue("background-cl");
              const fontWeight = style1.getPropertyValue("font-weight");
              // console.log({
              //   outer: parent.outerHTML,
              //   fontWeight,
              //   bold: fontWeight === 'bold',
              //   color: color,
              //   text: currentNode.textContent,
              // })
              temp.push(
                new TextRun({
                  bold: fontWeight === "bold" || fontWeight === "700",
                  color: color,
                  text: currentNode.textContent,
                })
              );
            }
            if (currentNodeName == this.mappableElements.BR) {
              temp.push(
                new TextRun({
                  break: true,
                })
              );
            }
          }
          docxObjects.push(
            new Paragraph({
              indent: {
                start: convertMillimetersToTwip(20.1),
              },
              spacing: {
                after: 180,
              },
              children: temp,
            })
          );
          //this.mapParagraph(i);
          continue;
        }
        case this.mappableElements.UL: {
          //const p = new Paragraph({});
          const iterator = Doc.createNodeIterator(
            firstLevelNode,
            NodeFilter.SHOW_ALL
          );
          let currentNode = iterator.nextNode();
          let temp = [];
          let stack = [];
          console.log("stack");
          while ((currentNode = iterator.nextNode())) {
            const currentNodeName = currentNode.nodeName.toLowerCase();
            if ([this.mappableElements.LI].includes(currentNodeName)) {
              const parent = currentNode.parentElement;
              let style1 = Window.getComputedStyle(parent);
              const color = style1.getPropertyValue("color");
              const foregroundColor = style1.getPropertyValue("background-cl");
              const fontWeight = style1.getPropertyValue("font-weight");
              // console.log({
              //   outer: parent.outerHTML,
              //   fontWeight,
              //   bold: fontWeight === 'bold',
              //   color: color,
              //   text: currentNode.textContent,
              // })
              temp.push(
                new Paragraph({
                  indent: {
                    start: convertMillimetersToTwip(20.1),
                  },
                  numbering: {
                    reference: "bullet-points",
                    level: 0,
                  },
                  children: [
                    new TextRun({
                      bold: fontWeight === "bold" || fontWeight === "700",
                      color: color,
                      text: currentNode.textContent,
                    }),
                  ],
                })
              );
            }
            if (currentNodeName == this.mappableElements.BR) {
              temp.push(
                new Paragraph({
                  indent: {
                    start: convertMillimetersToTwip(20.1),
                  },
                  children: [
                    new TextRun({
                      break: true,
                    }),
                  ],
                })
              );
            }
          }
          docxObjects.push(...temp);
          //this.mapParagraph(i);
          continue;
        }
        case this.mappableElements.OL: {
          //  this.mapUnOrderedList(i);
          continue;
        }
        case this.mappableElements.TABLE: {
          // this.mapTable(i);
          continue;
        }
      }
      //   console.log(`END------------->${nodeName}`);
    }

    return docxObjects;

    // let currentNode;
    // let p;
    // while ((currentNode = iterator.nextNode())) {
    //   if (currentNode.nodeName === "p" && p) {
    //     docxObjects.push(p);
    //     p = new Paragraph({});
    //   }
    //   if (currentNode.nodeName === "p" && !p) {
    //     p = new Paragraph({});
    //   }

    //   //  console.log({ currentNode });
    // }
    // console.log("_______________________end");

    // // for (const element of dom.window.document.body.children) {
    // //   const mapped = this.mapElement(element);
    // //   if (mapped) docxObjects.push(mapped);
    // // }

    // // console.log({ docxObjects });
    // return docxObjects;
  }

  mapElement(element) {
    switch (element.localName) {
      case this.mappableElements.P: {
        return this.mapParagraph(element);
      }
      case this.mappableElements.UL: {
        return this.mapUnOrderedList(element);
      }
      case this.mappableElements.OL: {
        return this.mapOrderedList(element);
      }
      case this.mappableElements.BR: {
        return this.mapBreakLine(element);
      }
      case this.mappableElements.SPAN: {
        return this.mapSpan(element);
      }
      case this.mappableElements.TABLE: {
        return this.mapTable(element);
      }
      default: {
        console.warn(`Not mapped:${element.tagName}`);
      }
    }
  }

  mapTable(element) {}

  mapTableCol(element) {}

  mapTableRow(element) {}

  mapParagraph(element) {
    const p = new Paragraph({
      size: "11pt",
      break: true,
      //text: htmlToText(element.outerHTML),
    });
    const innerHtml = element.innerHtml;
    for (const iterator of object) {
    }
  }

  mapSpan(element) {}

  mapOrderedList(element) {}
  mapUnOrderedList(element) {}

  mapBreakLine(element) {}

  mapListItem(element) {}

  applyStyles(styles) {}
}

module.exports = DOCXPublisher;

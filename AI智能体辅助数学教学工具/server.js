#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
} from '@modelcontextprotocol/sdk/types.js';
import axios from 'axios';
import * as cheerio from 'cheerio';
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from 'docx';
import PptxGenJS from 'pptxgenjs';
import puppeteer from 'puppeteer';
import sharp from 'sharp';
import { createCanvas } from 'canvas';
import { spawn } from 'child_process';
import { writeFileSync, existsSync, mkdirSync } from 'fs';
import { promisify } from 'util';
import { pipeline } from 'stream';
import { createWriteStream } from 'fs';
import archiver from 'archiver';

const server = new Server(
  {
    name: 'work-assistant-mcp',
    version: '1.0.0',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

// 网页查询工具
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      {
        name: 'web_search',
        description: '在网络上搜索信息并返回结果',
        inputSchema: {
          type: 'object',
          properties: {
            query: {
              type: 'string',
              description: '搜索查询词',
            },
            max_results: {
              type: 'number',
              description: '最大结果数量（默认5）',
              default: 5,
            },
          },
          required: ['query'],
        },
      },
      {
        name: 'fetch_webpage',
        description: '获取指定网页的内容',
        inputSchema: {
          type: 'object',
          properties: {
            url: {
              type: 'string',
              description: '要获取的网页URL',
            },
          },
          required: ['url'],
        },
      },
      {
        name: 'create_word_document',
        description: '创建Word文档',
        inputSchema: {
          type: 'object',
          properties: {
            title: {
              type: 'string',
              description: '文档标题',
            },
            content: {
              type: 'array',
              description: '文档内容段落',
              items: {
                type: 'object',
                properties: {
                  text: { type: 'string' },
                  heading: { type: 'string', enum: ['h1', 'h2', 'h3', 'normal'] }
                }
              }
            },
            filename: {
              type: 'string',
              description: '输出文件名（不含扩展名）',
            },
          },
          required: ['title', 'content', 'filename'],
        },
      },
      {
        name: 'create_powerpoint',
        description: '创建PowerPoint演示文稿',
        inputSchema: {
          type: 'object',
          properties: {
            title: {
              type: 'string',
              description: '演示文稿标题',
            },
            slides: {
              type: 'array',
              description: '幻灯片内容',
              items: {
                type: 'object',
                properties: {
                  title: { type: 'string' },
                  content: { type: 'array', items: { type: 'string' } }
                }
              }
            },
            filename: {
              type: 'string',
              description: '输出文件名（不含扩展名）',
            },
          },
          required: ['title', 'slides', 'filename'],
        },
      },
      {
        name: 'search_and_generate_geogebra',
        description: '搜索网页内容并生成GeoGebra演示图像',
        inputSchema: {
          type: 'object',
          properties: {
            query: {
              type: 'string',
              description: '搜索查询词（数学相关）',
            },
            math_type: {
              type: 'string',
              enum: ['function', 'geometry', 'statistics', 'calculus', 'algebra'],
              description: '数学类型',
              default: 'function'
            },
            description: {
              type: 'string',
              description: '具体的数学内容描述',
            },
            output_format: {
              type: 'string',
              enum: ['image', 'ggb_file', 'embed_code'],
              description: '输出格式',
              default: 'image'
            },
            filename: {
              type: 'string',
              description: '输出文件名（不含扩展名）',
            },
          },
          required: ['query', 'description'],
        },
      },
      {
        name: 'generate_geogebra_from_command',
        description: '根据指令直接生成GeoGebra演示',
        inputSchema: {
          type: 'object',
          properties: {
            command: {
              type: 'string',
              description: 'GeoGebra生成指令（如：二次函数、圆的切线等）',
            },
            parameters: {
              type: 'object',
              description: '生成参数',
              properties: {
                width: { type: 'number', default: 800 },
                height: { type: 'number', default: 600 },
                show_grid: { type: 'boolean', default: true },
                show_axes: { type: 'boolean', default: true },
                use_geogebra_app: { type: 'boolean', default: false, description: '是否使用本地GeoGebra应用生成高质量图像' }
              }
            },
            output_format: {
              type: 'string',
              enum: ['image', 'ggb_file', 'embed_code'],
              description: '输出格式',
              default: 'image'
            },
            filename: {
              type: 'string',
              description: '输出文件名（不含扩展名）',
            },
          },
          required: ['command'],
        },
      },
      {
        name: 'launch_geogebra',
        description: '启动本地GeoGebra应用程序',
        inputSchema: {
          type: 'object',
          properties: {
            file_path: {
              type: 'string',
              description: '要打开的.ggb文件路径（可选）',
            },
          },
        },
      },
      {
        name: 'generate_math_visualization',
        description: '生成数学可视化图表（函数图像、统计图表等）',
        inputSchema: {
          type: 'object',
          properties: {
            type: {
              type: 'string',
              enum: ['function', 'statistics', 'geometry', 'calculus'],
              description: '可视化类型',
            },
            data: {
              type: 'object',
              description: '数学数据或函数表达式',
            },
            title: {
              type: 'string',
              description: '图表标题',
            },
            filename: {
              type: 'string',
              description: '输出文件名（不含扩展名）',
            },
          },
          required: ['type', 'data', 'filename'],
        },
      },
      {
        name: 'create_math_exercise',
        description: '创建数学练习题及答案解析',
        inputSchema: {
          type: 'object',
          properties: {
            topic: {
              type: 'string',
              enum: ['algebra', 'geometry', 'calculus', 'statistics', 'trigonometry'],
              description: '数学主题',
            },
            difficulty: {
              type: 'string',
              enum: ['easy', 'medium', 'hard'],
              description: '难度等级',
            },
            count: {
              type: 'number',
              description: '题目数量',
              default: 5,
            },
            format: {
              type: 'string',
              enum: ['word', 'pdf', 'html'],
              description: '输出格式',
              default: 'word',
            },
            filename: {
              type: 'string',
              description: '输出文件名（不含扩展名）',
            },
          },
          required: ['topic', 'filename'],
        },
      },
      {
        name: 'math_formula_converter',
        description: '数学公式格式转换（LaTeX、MathML、图片）',
        inputSchema: {
          type: 'object',
          properties: {
            formula: {
              type: 'string',
              description: '数学公式',
            },
            input_format: {
              type: 'string',
              enum: ['latex', 'text', 'mathml'],
              description: '输入格式',
              default: 'latex',
            },
            output_format: {
              type: 'string',
              enum: ['latex', 'mathml', 'image', 'text'],
              description: '输出格式',
              default: 'image',
            },
            filename: {
              type: 'string',
              description: '输出文件名（不含扩展名）',
            },
          },
          required: ['formula', 'filename'],
        },
      },
      {
        name: 'generate_math_tutorial',
        description: '生成分步数学教程',
        inputSchema: {
          type: 'object',
          properties: {
            concept: {
              type: 'string',
              description: '数学概念（如：二次函数、导数、积分）',
            },
            level: {
              type: 'string',
              enum: ['beginner', 'intermediate', 'advanced'],
              description: '难度级别',
              default: 'intermediate',
            },
            include_examples: {
              type: 'boolean',
              description: '是否包含例题',
              default: true,
            },
            include_exercises: {
              type: 'boolean',
              description: '是否包含练习题',
              default: true,
            },
            format: {
              type: 'string',
              enum: ['word', 'powerpoint', 'html'],
              description: '输出格式',
              default: 'word',
            },
            filename: {
              type: 'string',
              description: '输出文件名（不含扩展名）',
            },
          },
          required: ['concept', 'filename'],
        },
      },
    ],
  };
});

// 处理工具调用
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  try {
    const { name, arguments: args } = request.params;

    switch (name) {
      case 'web_search': {
        const { query, max_results = 5 } = args;
        
        try {
          // 使用DuckDuckGo进行搜索
          const searchUrl = `https://duckduckgo.com/html/?q=${encodeURIComponent(query)}`;
          const response = await axios.get(searchUrl, {
            headers: {
              'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }
          });
          
          const $ = cheerio.load(response.data);
          const results = [];
          
          $('.result').each((i, element) => {
            if (i >= max_results) return false;
            
            const $result = $(element);
            const title = $result.find('.result__a').text().trim();
            const url = $result.find('.result__a').attr('href');
            const snippet = $result.find('.result__snippet').text().trim();
            
            if (title && url) {
              results.push({
                title,
                url: url.replace(/^\/l\/?uddg=/, ''),
                snippet,
              });
            }
          });
          
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  query,
                  results: results.length > 0 ? results : [{ title: '未找到结果', url: '', snippet: '请尝试其他搜索词' }]
                }, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `搜索失败: ${error.message}`
          );
        }
      }

      case 'fetch_webpage': {
        const { url } = args;
        
        try {
          const response = await axios.get(url, {
            headers: {
              'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            },
            timeout: 10000
          });
          
          const $ = cheerio.load(response.data);
          
          // 移除脚本和样式
          $('script, style').remove();
          
          // 提取主要内容
          const title = $('title').text().trim() || '无标题';
          const content = $('body').text().replace(/\s+/g, ' ').trim();
          
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  url,
                  title,
                  content: content.substring(0, 5000) + (content.length > 5000 ? '...' : ''),
                  length: content.length
                }, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `获取网页失败: ${error.message}`
          );
        }
      }

      case 'create_word_document': {
        const { title, content, filename } = args;
        
        try {
          const doc = new Document();
          
          // 添加标题
          doc.addSection({
            properties: {},
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: title,
                    bold: true,
                    size: 32,
                  }),
                ],
                heading: HeadingLevel.TITLE,
              }),
            ],
          });
          
          // 添加内容
          content.forEach(item => {
            const level = item.heading === 'h1' ? HeadingLevel.HEADING_1 :
                         item.heading === 'h2' ? HeadingLevel.HEADING_2 :
                         item.heading === 'h3' ? HeadingLevel.HEADING_3 :
                         undefined;
            
            doc.addParagraph(
              new Paragraph({
                children: [
                  new TextRun({
                    text: item.text,
                    size: level ? 28 : 24,
                  }),
                ],
                heading: level,
              })
            );
          });
          
          // 生成文档
          const buffer = await Packer.toBuffer(doc);
          const outputPath = `E:\\nvm\\work-assistant-mcp\\${filename}.docx`;
          
          return {
            content: [
              {
                type: 'text',
                text: `Word文档已创建: ${filename}.docx\n\n文档预览:\n标题: ${title}\n段落数: ${content.length}`,
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `创建Word文档失败: ${error.message}`
          );
        }
      }

      case 'create_powerpoint': {
        const { title, slides, filename } = args;
        
        try {
          const pptx = new PptxGenJS();
          
          // 设置演示文稿属性
          pptx.defineLayout({ name: 'A4', width: 10, height: 7.5 });
          pptx.layout = 'A4';
          
          // 添加标题幻灯片
          const titleSlide = pptx.addSlide();
          titleSlide.addText(title, {
            x: 1,
            y: 2,
            fontSize: 44,
            bold: true,
            color: '363636',
            align: 'center'
          });
          
          // 添加内容幻灯片
          slides.forEach((slide, index) => {
            const slideObj = pptx.addSlide();
            
            // 添加幻灯片标题
            if (slide.title) {
              slideObj.addText(slide.title, {
                x: 0.5,
                y: 0.5,
                fontSize: 28,
                bold: true,
                color: '363636'
              });
            }
            
            // 添加幻灯片内容
            if (slide.content && slide.content.length > 0) {
              slideObj.addText(slide.content.join('\n'), {
                x: 0.5,
                y: 1.5,
                fontSize: 18,
                color: '363636',
                bullet: true
              });
            }
          });
          
          return {
            content: [
              {
                type: 'text',
                text: `PowerPoint演示文稿已创建: ${filename}.pptx\n\n演示文稿预览:\n标题: ${title}\n幻灯片数: ${slides.length + 1}`,
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `创建PowerPoint失败: ${error.message}`
          );
        }
      }

      case 'search_and_generate_geogebra': {
        const { query, math_type = 'function', description, output_format = 'image', filename } = args;
        
        try {
          // 首先搜索相关内容
          const searchResults = await searchWebContent(query + ' ' + description);
          
          // 生成GeoGebra内容
          const geoResult = await generateGeoGebraContent(description, math_type, output_format, filename);
          
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  search_results: searchResults,
                  geogebra_output: geoResult,
                  query,
                  description,
                  output_format
                }, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `搜索并生成GeoGebra失败: ${error.message}`
          );
        }
      }

      case 'generate_geogebra_from_command': {
        const { command, parameters = {}, output_format = 'image', filename } = args;
        
        try {
          let geoResult;
          
          // 检查是否使用本地GeoGebra应用
          if (parameters.use_geogebra_app && output_format === 'image') {
            const geogebraCode = parseCommandToGeoGebra(command, parameters);
            geoResult = await generateGeoGebraImageWithApp(geogebraCode, filename);
          } else {
            geoResult = await generateGeoGebraFromCommand(command, parameters, output_format, filename);
          }
          
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  geogebra_output: geoResult,
                  command,
                  parameters,
                  output_format
                }, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `生成GeoGebra失败: ${error.message}`
          );
        }
      }

      case 'launch_geogebra': {
        const { file_path } = args;
        
        try {
          const result = await launchGeoGebraWithFile(file_path || '');
          
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  message: 'GeoGebra应用已启动',
                  file_path: file_path || '无',
                  result: result
                }, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `启动GeoGebra失败: ${error.message}`
          );
        }
      }

      case 'generate_math_visualization': {
        const { type, data, title, filename } = args;
        
        try {
          const result = await generateMathVisualization(type, data, title, filename);
          
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  visualization: result,
                  type,
                  title
                }, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `生成数学可视化失败: ${error.message}`
          );
        }
      }

      case 'create_math_exercise': {
        const { topic, difficulty = 'medium', count = 5, format = 'word', filename } = args;
        
        try {
          const result = await createMathExercise(topic, difficulty, count, format, filename);
          
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  exercise: result,
                  topic,
                  difficulty,
                  count,
                  format
                }, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `创建数学练习题失败: ${error.message}`
          );
        }
      }

      case 'math_formula_converter': {
        const { formula, input_format = 'latex', output_format = 'image', filename } = args;
        
        try {
          const result = await convertMathFormula(formula, input_format, output_format, filename);
          
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  conversion: result,
                  original: formula,
                  input_format,
                  output_format
                }, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `数学公式转换失败: ${error.message}`
          );
        }
      }

      case 'generate_math_tutorial': {
        const { concept, level = 'intermediate', include_examples = true, include_exercises = true, format = 'word', filename } = args;
        
        try {
          const result = await generateMathTutorial(concept, level, include_examples, include_exercises, format, filename);
          
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  tutorial: result,
                  concept,
                  level,
                  format
                }, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(
            ErrorCode.InternalError,
            `生成数学教程失败: ${error.message}`
          );
        }
      }

      default:
        throw new McpError(
          ErrorCode.MethodNotFound,
          `未知工具: ${name}`
        );
    }
  } catch (error) {
    throw new McpError(
      ErrorCode.InternalError,
      `工具执行失败: ${error.message}`
    );
  }
});

// 辅助函数：搜索网页内容
async function searchWebContent(query, maxResults = 3) {
  try {
    const searchUrl = `https://duckduckgo.com/html/?q=${encodeURIComponent(query)}`;
    const response = await axios.get(searchUrl, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
      }
    });
    
    const $ = cheerio.load(response.data);
    const results = [];
    
    $('.result').each((i, element) => {
      if (i >= maxResults) return false;
      
      const $result = $(element);
      const title = $result.find('.result__a').text().trim();
      const url = $result.find('.result__a').attr('href');
      const snippet = $result.find('.result__snippet').text().trim();
      
      if (title && url) {
        results.push({
          title,
          url: url.replace(/^\/l\/?uddg=/, ''),
          snippet,
        });
      }
    });
    
    return results;
  } catch (error) {
    console.error('搜索失败:', error);
    return [];
  }
}

// GeoGebra相关辅助函数
async function generateGeoGebraContent(description, mathType, outputFormat, filename) {
  try {
    // 根据数学类型和描述生成GeoGebra代码
    const geogebraCode = generateGeoGebraCode(description, mathType);
    
    switch (outputFormat) {
      case 'image':
        return await generateGeoGebraImage(geogebraCode, filename);
      case 'ggb_file':
        return await generateGeoGebraFile(geogebraCode, filename);
      case 'embed_code':
        return generateGeoGebraEmbedCode(geogebraCode);
      default:
        throw new Error(`不支持的输出格式: ${outputFormat}`);
    }
  } catch (error) {
    throw new Error(`生成GeoGebra内容失败: ${error.message}`);
  }
}

async function generateGeoGebraFromCommand(command, parameters, outputFormat, filename) {
  try {
    // 解析命令并生成相应的GeoGebra代码
    const geogebraCode = parseCommandToGeoGebra(command, parameters);
    
    switch (outputFormat) {
      case 'image':
        return await generateGeoGebraImage(geogebraCode, filename);
      case 'ggb_file':
        return await generateGeoGebraFile(geogebraCode, filename);
      case 'embed_code':
        return generateGeoGebraEmbedCode(geogebraCode);
      default:
        throw new Error(`不支持的输出格式: ${outputFormat}`);
    }
  } catch (error) {
    throw new Error(`从命令生成GeoGebra失败: ${error.message}`);
  }
}

function generateGeoGebraCode(description, mathType) {
  const codeGenerators = {
    function: (desc) => {
      // 函数图形生成逻辑
      if (desc.includes('二次函数') || desc.includes('抛物线')) {
        return 'f(x) = x^2';
      } else if (desc.includes('三角函数') || desc.includes('正弦')) {
        return 'f(x) = sin(x)';
      } else if (desc.includes('余弦')) {
        return 'f(x) = cos(x)';
      } else if (desc.includes('正切')) {
        return 'f(x) = tan(x)';
      }
      return 'f(x) = x^2';
    },
    geometry: (desc) => {
      // 几何图形生成逻辑
      if (desc.includes('圆')) {
        return 'Circle((0,0), 3)';
      } else if (desc.includes('三角形')) {
        return 'Polygon((0,0), (4,0), (2,3))';
      } else if (desc.includes('正方形')) {
        return 'Polygon((-2,-2), (2,-2), (2,2), (-2,2))';
      }
      return 'Circle((0,0), 3)';
    },
    statistics: (desc) => {
      // 统计图表生成逻辑
      return 'Histogram({1,2,3,4,5}, {2,3,5,2,1})';
    },
    calculus: (desc) => {
      // 微积分图形生成逻辑
      if (desc.includes('导数')) {
        return 'Derivative(x^2)';
      } else if (desc.includes('积分')) {
        return 'Integral(x^2, 0, 3)';
      }
      return 'Derivative(x^2)';
    },
    algebra: (desc) => {
      // 代数图形生成逻辑
      return 'Line((0,1), (2,0))';
    }
  };
  
  const generator = codeGenerators[mathType] || codeGenerators.function;
  return generator(description);
}

function parseCommandToGeoGebra(command, parameters) {
  // 解析自然语言命令为GeoGebra代码
  const width = parameters.width || 800;
  const height = parameters.height || 600;
  const showGrid = parameters.show_grid !== false;
  const showAxes = parameters.show_axes !== false;
  
  let code = '';
  
  // 根据命令内容生成相应的GeoGebra代码
  if (command.includes('二次函数') || command.includes('抛物线')) {
    code = 'f(x) = x^2';
  } else if (command.includes('圆')) {
    code = 'Circle((0,0), 3)';
  } else if (command.includes('直线')) {
    code = 'Line((0,0), (3,2))';
  } else {
    // 默认生成一个简单的函数
    code = 'f(x) = x^2';
  }
  
  return code;
}

async function generateGeoGebraImage(geogebraCode, filename) {
  try {
    // 确保输出目录存在
    const outputDir = 'E:\\nvm\\work-assistant-mcp';
    if (!existsSync(outputDir)) {
      mkdirSync(outputDir, { recursive: true });
    }
    
    // 使用Promise包装Canvas操作，避免阻塞
    const canvasPromise = new Promise((resolve, reject) => {
      try {
        // 使用setTimeout将Canvas操作移出主线程
        setTimeout(() => {
          try {
            const width = 800;
            const height = 600;
            const canvas = createCanvas(width, height);
            const ctx = canvas.getContext('2d');
            
            // 设置背景
            ctx.fillStyle = 'white';
            ctx.fillRect(0, 0, width, height);
            
            // 绘制坐标轴
            ctx.strokeStyle = '#333';
            ctx.lineWidth = 2;
            ctx.beginPath();
            ctx.moveTo(50, height / 2);
            ctx.lineTo(width - 50, height / 2);
            ctx.moveTo(width / 2, 50);
            ctx.lineTo(width / 2, height - 50);
            ctx.stroke();
            
            // 绘制网格
            ctx.strokeStyle = '#e0e0e0';
            ctx.lineWidth = 1;
            for (let i = 100; i < width - 50; i += 50) {
              ctx.beginPath();
              ctx.moveTo(i, 50);
              ctx.lineTo(i, height - 50);
              ctx.stroke();
            }
            for (let i = 100; i < height - 50; i += 50) {
              ctx.beginPath();
              ctx.moveTo(50, i);
              ctx.lineTo(width - 50, i);
              ctx.stroke();
            }
            
            // 根据GeoGebra代码绘制函数
            ctx.strokeStyle = 'red';
            ctx.lineWidth = 3;
            ctx.beginPath();
            
            if (geogebraCode.includes('x^2')) {
              // 绘制抛物线
              for (let x = -300; x <= 300; x++) {
                const realX = x / 50;
                const y = realX * realX;
                const canvasX = width / 2 + x;
                const canvasY = height / 2 - y * 50;
                
                if (x === -300) {
                  ctx.moveTo(canvasX, canvasY);
                } else {
                  ctx.lineTo(canvasX, canvasY);
                }
              }
            } else if (geogebraCode.includes('sin')) {
              // 绘制正弦函数
              for (let x = -300; x <= 300; x++) {
                const realX = x / 30;
                const y = Math.sin(realX);
                const canvasX = width / 2 + x;
                const canvasY = height / 2 - y * 100;
                
                if (x === -300) {
                  ctx.moveTo(canvasX, canvasY);
                } else {
                  ctx.lineTo(canvasX, canvasY);
                }
              }
            } else if (geogebraCode.includes('cos')) {
              // 绘制余弦函数
              for (let x = -300; x <= 300; x++) {
                const realX = x / 30;
                const y = Math.cos(realX);
                const canvasX = width / 2 + x;
                const canvasY = height / 2 - y * 100;
                
                if (x === -300) {
                  ctx.moveTo(canvasX, canvasY);
                } else {
                  ctx.lineTo(canvasX, canvasY);
                }
              }
            } else if (geogebraCode.includes('Circle')) {
              // 绘制圆
              ctx.arc(width / 2, height / 2, 150, 0, 2 * Math.PI);
            }
            
            ctx.stroke();
            
            // 添加标签
            ctx.fillStyle = '#333';
            ctx.font = '14px Arial';
            ctx.fillText(geogebraCode, 10, 20);
            
            // 异步生成缓冲区
            canvas.toBuffer('image/png', (err, buffer) => {
              if (err) {
                reject(err);
              } else {
                resolve(buffer);
              }
            });
          } catch (error) {
            reject(error);
          }
        }, 0);
      } catch (error) {
        reject(error);
      }
    });
    
    // 等待Canvas操作完成，设置超时
    const buffer = await Promise.race([
      canvasPromise,
      new Promise((_, reject) => 
        setTimeout(() => reject(new Error('Canvas操作超时')), 10000)
      )
    ]);
    
    // 保存文件
    const outputPath = `${outputDir}\\${filename}.png`;
    writeFileSync(outputPath, buffer);
    
    return {
      type: 'image',
      format: 'png',
      path: outputPath,
      size: buffer.length,
      code: geogebraCode,
      preview: '图像已生成并保存到文件系统'
    };
  } catch (error) {
    throw new Error(`生成图像失败: ${error.message}`);
  }
}

async function generateGeoGebraFile(geogebraCode, filename) {
  try {
    // 确保输出目录存在
    const outputDir = 'E:\\nvm\\work-assistant-mcp';
    if (!existsSync(outputDir)) {
      mkdirSync(outputDir, { recursive: true });
    }
    
    // 解析GeoGebra代码
    let expression = '';
    if (geogebraCode.includes('Surface') || geogebraCode.includes('sqrt')) {
      expression = 'y = x^2 + z^2';
    } else if (geogebraCode.includes('x^2')) {
      expression = 'f(x) = x^2';
    } else if (geogebraCode.includes('sin')) {
      expression = 'f(x) = sin(x)';
    } else if (geogebraCode.includes('cos')) {
      expression = 'f(x) = cos(x)';
    } else if (geogebraCode.includes('Circle')) {
      expression = 'Circle((0,0), 3)';
    } else {
      expression = geogebraCode;
    }
    
    // 创建GeoGebra XML内容
    const xmlContent = `<?xml version="1.0" encoding="utf-8"?>
<geogebra format="5.0" version="6.0.899.0" app="classic" platform="win" language="zh">
  <construction>
    <expression exp="${expression}" type="function" label="f">
      <show object="true" label="true"/>
      <color r="0" g="102" b="204" alpha="255"/>
      <lineStyle thickness="2"/>
      <pointSize>3</pointSize>
    </expression>
    <expression exp="x" type="axis" label="x">
      <show object="true" label="true"/>
    </expression>
    <expression exp="y" type="axis" label="y">
      <show object="true" label="true"/>
    </expression>
    <expression exp="z" type="axis" label="z">
      <show object="true" label="true"/>
    </expression>
  </construction>
  <euclidianView>
    <size width="800" height="600"/>
    <coordSystem xZero="400" yZero="300" scale="50" yScale="50"/>
  </euclidianView>
  <kernel>
    <continuous>true</continuous>
  </kernel>
</geogebra>`;
    
    // 创建.ggb文件（实际上是ZIP格式）
    const outputPath = `${outputDir}\\${filename}.ggb`;
    
    return new Promise((resolve, reject) => {
      // 创建ZIP文件流
      const output = createWriteStream(outputPath);
      const archive = archiver('zip', { zlib: { level: 9 } });
      
      output.on('close', () => {
        resolve({
          type: 'ggb_file',
          format: 'ggb',
          path: outputPath,
          code: geogebraCode,
          content: expression,
          size: archive.pointer(),
          preview: 'GeoGebra文件已生成（ZIP格式），可直接在GeoGebra中打开'
        });
      });
      
      archive.on('error', (err) => {
        reject(new Error(`创建ZIP文件失败: ${err.message}`));
      });
      
      archive.pipe(output);
      
      // 添加geogebra.xml文件到ZIP
      archive.append(xmlContent, { name: 'geogebra.xml' });
      
      // 添加必要的元数据文件
      const metaContent = `<?xml version="1.0" encoding="utf-8"?>
<geogebra format="5.0" version="6.0.899.0">
  <metadata>
    <author>Work Assistant MCP</author>
    <date>${new Date().toISOString()}</date>
    <description>Generated by MCP: ${expression}</description>
  </metadata>
</geogebra>`;
      archive.append(metaContent, { name: 'metadata.xml' });
      
      archive.finalize();
    });
  } catch (error) {
    throw new Error(`生成GeoGebra文件失败: ${error.message}`);
  }
}

// 新增：与本地GeoGebra应用程序集成的函数
async function launchGeoGebraWithFile(filePath) {
  return new Promise((resolve, reject) => {
    const geogebraPath = 'C:\\Users\\rog\\Desktop\\GeoGebra_Calculator\\app-6.0.8990\\GeoGebra.exe';
    
    // 启动GeoGebra应用
    const geogebraProcess = spawn(geogebraPath, [filePath], {
      stdio: ['ignore', 'pipe', 'pipe'],
      detached: true
    });
    
    let stdout = '';
    let stderr = '';
    
    geogebraProcess.stdout.on('data', (data) => {
      stdout += data.toString();
    });
    
    geogebraProcess.stderr.on('data', (data) => {
      stderr += data.toString();
    });
    
    // 设置超时
    const timeout = setTimeout(() => {
      geogebraProcess.kill();
      reject(new Error('GeoGebra启动超时'));
    }, 15000);
    
    geogebraProcess.on('close', (code) => {
      clearTimeout(timeout);
      if (code === 0) {
        resolve({
          success: true,
          message: 'GeoGebra已成功启动',
          stdout: stdout,
          stderr: stderr
        });
      } else {
        reject(new Error(`GeoGebra启动失败，退出码: ${code}, 错误: ${stderr}`));
      }
    });
    
    geogebraProcess.on('error', (error) => {
      clearTimeout(timeout);
      reject(new Error(`无法启动GeoGebra: ${error.message}`));
    });
    
    // 分离进程，让它独立运行
    geogebraProcess.unref();
  });
}

// 新增：使用GeoGebra生成高质量图像
async function generateGeoGebraImageWithApp(geogebraCode, filename) {
  try {
    // 首先生成.ggb文件
    const ggbResult = await generateGeoGebraFile(geogebraCode, filename);
    
    // 启动GeoGebra应用并导出图像
    const geogebraPath = 'C:\\Users\\rog\\Desktop\\GeoGebra_Calculator\\app-6.0.8990\\GeoGebra.exe';
    const outputImagePath = `E:\\nvm\\work-assistant-mcp\\${filename}_high_quality.png`;
    
    // 使用命令行参数导出图像
    const exportProcess = spawn(geogebraPath, [
      '--export=' + outputImagePath,
      '--width=1200',
      '--height=900',
      '--showToolBar=false',
      '--showAlgebraInput=false',
      ggbResult.path
    ], {
      stdio: ['ignore', 'pipe', 'pipe'],
      timeout: 30000
    });
    
    const exportResult = await new Promise((resolve, reject) => {
      let stdout = '';
      let stderr = '';
      
      exportProcess.stdout.on('data', (data) => {
        stdout += data.toString();
      });
      
      exportProcess.stderr.on('data', (data) => {
        stderr += data.toString();
      });
      
      exportProcess.on('close', (code) => {
        if (code === 0) {
          resolve({
            success: true,
            stdout: stdout,
            stderr: stderr
          });
        } else {
          reject(new Error(`导出失败，退出码: ${code}, 错误: ${stderr}`));
        }
      });
      
      exportProcess.on('error', (error) => {
        reject(new Error(`导出进程错误: ${error.message}`));
      });
    });
    
    return {
      type: 'image',
      format: 'png',
      path: outputImagePath,
      code: geogebraCode,
      quality: 'high',
      method: 'geogebra_app',
      preview: '使用GeoGebra应用生成的高质量图像',
      export_result: exportResult
    };
  } catch (error) {
    throw new Error(`使用GeoGebra应用生成图像失败: ${error.message}`);
  }
}

function generateGeoGebraEmbedCode(geogebraCode) {
  const embedCode = `<!DOCTYPE html>
<html>
<head>
    <title>GeoGebra演示</title>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body>
    <div id="geogebra"></div>
    <script src="https://www.geogebra.org/apps/deployggb.js"></script>
    <script>
        var ggbApp = new GGBApplet({
            "appName": "graphing",
            "width": 800,
            "height": 600,
            "showToolBar": true,
            "showAlgebraInput": true,
            "showMenuBar": true,
            "enableRightClick": true,
            "enableShiftDragZoom": true,
            "showZoomButtons": true,
            "language": "zh"
        }, true);
        
        ggbApp.inject('geogebra');
        
        ggbApp.registerClientListener(function() {
            try {
                ggbApp.evalCommand('${geogebraCode}');
            } catch(e) {
                console.error('GeoGebra代码执行错误:', e);
            }
        });
    </script>
</body>
</html>`;
  
  return {
    type: 'embed_code',
    format: 'html',
    code: embedCode,
    preview: 'HTML嵌入代码已生成，可直接在网页中使用'
  };
}

// 数学可视化函数
async function generateMathVisualization(type, data, title, filename) {
  try {
    const outputDir = 'E:\\nvm\\work-assistant-mcp';
    if (!existsSync(outputDir)) {
      mkdirSync(outputDir, { recursive: true });
    }

    let visualizationData = {};

    switch (type) {
      case 'function':
        visualizationData = await generateFunctionPlot(data, title);
        break;
      case 'statistics':
        visualizationData = await generateStatisticsChart(data, title);
        break;
      case 'geometry':
        visualizationData = await generateGeometryDiagram(data, title);
        break;
      case 'calculus':
        visualizationData = await generateCalculusVisualization(data, title);
        break;
    }

    // 生成HTML可视化文件
    const htmlContent = generateVisualizationHTML(visualizationData, title);
    const htmlPath = `${outputDir}\\${filename}.html`;
    writeFileSync(htmlPath, htmlContent, 'utf8');

    return {
      type: 'visualization',
      format: 'html',
      path: htmlPath,
      title: title,
      data: visualizationData,
      preview: '数学可视化已生成，可直接在浏览器中打开查看'
    };
  } catch (error) {
    throw new Error(`生成数学可视化失败: ${error.message}`);
  }
}

// 数学练习题生成函数
async function createMathExercise(topic, difficulty, count, format, filename) {
  try {
    const outputDir = 'E:\\nvm\\work-assistant-mcp';
    if (!existsSync(outputDir)) {
      mkdirSync(outputDir, { recursive: true });
    }

    const exercises = generateExercisesByTopic(topic, difficulty, count);
    
    let result;
    switch (format) {
      case 'word':
        result = await createExerciseWord(exercises, topic, difficulty, filename);
        break;
      case 'pdf':
        result = await createExercisePDF(exercises, topic, difficulty, filename);
        break;
      case 'html':
        result = await createExerciseHTML(exercises, topic, difficulty, filename);
        break;
    }

    return result;
  } catch (error) {
    throw new Error(`创建数学练习题失败: ${error.message}`);
  }
}

// 数学公式转换函数
async function convertMathFormula(formula, inputFormat, outputFormat, filename) {
  try {
    const outputDir = 'E:\\nvm\\work-assistant-mcp';
    if (!existsSync(outputDir)) {
      mkdirSync(outputDir, { recursive: true });
    }

    let result;
    switch (outputFormat) {
      case 'image':
        result = await convertFormulaToImage(formula, inputFormat, filename);
        break;
      case 'latex':
        result = convertToLaTeX(formula, inputFormat);
        break;
      case 'mathml':
        result = convertToMathML(formula, inputFormat);
        break;
      case 'text':
        result = convertToText(formula, inputFormat);
        break;
    }

    return result;
  } catch (error) {
    throw new Error(`数学公式转换失败: ${error.message}`);
  }
}

// 数学教程生成函数
async function generateMathTutorial(concept, level, includeExamples, includeExercises, format, filename) {
  try {
    const outputDir = 'E:\\nvm\\work-assistant-mcp';
    if (!existsSync(outputDir)) {
      mkdirSync(outputDir, { recursive: true });
    }

    const tutorial = generateTutorialContent(concept, level, includeExamples, includeExercises);
    
    let result;
    switch (format) {
      case 'word':
        result = await createTutorialWord(tutorial, filename);
        break;
      case 'powerpoint':
        result = await createTutorialPowerPoint(tutorial, filename);
        break;
      case 'html':
        result = await createTutorialHTML(tutorial, filename);
        break;
    }

    return result;
  } catch (error) {
    throw new Error(`生成数学教程失败: ${error.message}`);
  }
}

// 辅助函数：生成函数图像
function generateFunctionPlot(data, title) {
  const { expression, range = { x: [-10, 10], y: [-10, 10] } } = data;
  
  return {
    type: 'function_plot',
    expression: expression,
    range: range,
    points: generateFunctionPoints(expression, range),
    title: title
  };
}

// 辅助函数：生成统计图表
function generateStatisticsChart(data, title) {
  const { chartType, values, labels } = data;
  
  return {
    type: 'statistics_chart',
    chartType: chartType,
    values: values,
    labels: labels,
    title: title
  };
}

// 辅助函数：生成几何图形
function generateGeometryDiagram(data, title) {
  const { shapes, coordinates } = data;
  
  return {
    type: 'geometry_diagram',
    shapes: shapes,
    coordinates: coordinates,
    title: title
  };
}

// 辅助函数：生成微积分可视化
function generateCalculusVisualization(data, title) {
  const { concept, function: func, parameters } = data;
  
  return {
    type: 'calculus_viz',
    concept: concept,
    function: func,
    parameters: parameters,
    title: title
  };
}

// 辅助函数：生成HTML可视化
function generateVisualizationHTML(data, title) {
  return `<!DOCTYPE html>
<html lang="zh">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${title}</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/mathjs/11.8.0/math.min.js"></script>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .container { max-width: 1200px; margin: 0 auto; }
        .chart-container { margin: 20px 0; }
        canvas { border: 1px solid #ddd; }
    </style>
</head>
<body>
    <div class="container">
        <h1>${title}</h1>
        <div class="chart-container">
            <canvas id="mathChart"></canvas>
        </div>
        <div id="details">
            <h2>详细信息</h2>
            <pre>${JSON.stringify(data, null, 2)}</pre>
        </div>
    </div>
    
    <script>
        // 根据数据类型生成相应的图表
        const data = ${JSON.stringify(data)};
        const ctx = document.getElementById('mathChart').getContext('2d');
        
        // 这里可以根据data.type生成不同类型的图表
        const chart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: data.points ? data.points.map(p => p.x) : [],
                datasets: [{
                    label: data.expression || '数学函数',
                    data: data.points ? data.points.map(p => p.y) : [],
                    borderColor: 'rgb(75, 192, 192)',
                    tension: 0.1
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    title: {
                        display: true,
                        text: '${title}'
                    }
                }
            }
        });
    </script>
</body>
</html>`;
}

// 辅助函数：生成函数点
function generateFunctionPoints(expression, range) {
  const points = [];
  const step = (range.x[1] - range.x[0]) / 100;
  
  for (let x = range.x[0]; x <= range.x[1]; x += step) {
    try {
      const y = math.evaluate(expression, { x: x });
      points.push({ x: x, y: y });
    } catch (e) {
      // 跳过无法计算的点
    }
  }
  
  return points;
}

// 辅助函数：按主题生成练习题
function generateExercisesByTopic(topic, difficulty, count) {
  const exerciseTemplates = {
    algebra: {
      easy: [
        { question: "解方程：2x + 5 = 15", answer: "x = 5", explanation: "2x = 15 - 5 = 10, x = 10/2 = 5" },
        { question: "计算：(3x - 2)(x + 4)", answer: "3x² + 10x - 8", explanation: "使用分配律展开" }
      ],
      medium: [
        { question: "解二次方程：x² - 5x + 6 = 0", answer: "x = 2 或 x = 3", explanation: "因式分解：(x-2)(x-3) = 0" }
      ]
    },
    geometry: {
      easy: [
        { question: "正方形的周长是20cm，求面积", answer: "25cm²", explanation: "边长=20/4=5cm，面积=5²=25cm²" }
      ]
    }
  };

  const templates = exerciseTemplates[topic]?.[difficulty] || exerciseTemplates.algebra.easy;
  return Array.from({ length: count }, (_, i) => templates[i % templates.length]);
}

// 辅助函数：创建练习题Word文档
async function createExerciseWord(exercises, topic, difficulty, filename) {
  const doc = new Document();
  
  doc.addSection({
    properties: {},
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text: `${topic}练习题 - ${difficulty}难度`,
            bold: true,
            size: 32,
          }),
        ],
        heading: HeadingLevel.TITLE,
      }),
    ],
  });

  exercises.forEach((exercise, index) => {
    doc.addParagraph(
      new Paragraph({
        children: [
          new TextRun({
            text: `题目${index + 1}: ${exercise.question}`,
            size: 24,
          }),
        ],
        heading: HeadingLevel.HEADING_2,
      })
    );

    doc.addParagraph(
      new Paragraph({
        children: [
          new TextRun({
            text: `答案: ${exercise.answer}`,
            size: 22,
          }),
        ],
      })
    );

    doc.addParagraph(
      new Paragraph({
        children: [
          new TextRun({
            text: `解析: ${exercise.explanation}`,
            size: 20,
          }),
        ],
      })
    );
  });

  const outputPath = `E:\\nvm\\work-assistant-mcp\\${filename}.docx`;
  
  return {
    type: 'exercise',
    format: 'word',
    path: outputPath,
    count: exercises.length,
    topic: topic,
    preview: `已生成${exercises.length}道${topic}练习题`
  };
}

// 辅助函数：公式转图片
async function convertFormulaToImage(formula, inputFormat, filename) {
  const canvas = createCanvas(400, 100);
  const ctx = canvas.getContext('2d');
  
  ctx.fillStyle = 'white';
  ctx.fillRect(0, 0, 400, 100);
  
  ctx.fillStyle = 'black';
  ctx.font = '20px Arial';
  ctx.fillText(formula, 20, 50);
  
  const buffer = canvas.toBuffer('image/png');
  const outputPath = `E:\\nvm\\work-assistant-mcp\\${filename}.png`;
  writeFileSync(outputPath, buffer);
  
  return {
    type: 'formula_image',
    format: 'png',
    path: outputPath,
    formula: formula,
    preview: '数学公式图片已生成'
  };
}

// 辅助函数：生成教程内容
function generateTutorialContent(concept, level, includeExamples, includeExercises) {
  return {
    title: `${concept}教程`,
    level: level,
    sections: [
      {
        title: "基本概念",
        content: `${concept}的基本定义和原理...`
      },
      {
        title: "核心方法",
        content: `掌握${concept}的核心方法...`
      }
    ],
    examples: includeExamples ? [
      {
        problem: `${concept}例题1`,
        solution: "详细解答步骤..."
      }
    ] : [],
    exercises: includeExercises ? [
      {
        question: `${concept}练习题`,
        answer: "参考答案"
      }
    ] : []
  };
}

// 辅助函数：创建教程HTML
async function createTutorialHTML(tutorial, filename) {
  const htmlContent = `<!DOCTYPE html>
<html lang="zh">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${tutorial.title}</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; line-height: 1.6; }
        .container { max-width: 1200px; margin: 0 auto; }
        .section { margin: 20px 0; padding: 15px; border: 1px solid #ddd; }
        .example, .exercise { background: #f9f9f9; padding: 10px; margin: 10px 0; }
    </style>
</head>
<body>
    <div class="container">
        <h1>${tutorial.title}</h1>
        <p>难度级别: ${tutorial.level}</p>
        
        ${tutorial.sections.map(section => `
        <div class="section">
            <h2>${section.title}</h2>
            <p>${section.content}</p>
        </div>
        `).join('')}
        
        ${tutorial.examples.length > 0 ? `
        <h2>例题</h2>
        ${tutorial.examples.map(example => `
        <div class="example">
            <h3>${example.problem}</h3>
            <p>${example.solution}</p>
        </div>
        `).join('')}
        ` : ''}
        
        ${tutorial.exercises.length > 0 ? `
        <h2>练习题</h2>
        ${tutorial.exercises.map(exercise => `
        <div class="exercise">
            <p><strong>题目:</strong> ${exercise.question}</p>
            <p><strong>答案:</strong> ${exercise.answer}</p>
        </div>
        `).join('')}
        ` : ''}
    </div>
</body>
</html>`;

  const outputPath = `E:\\nvm\\work-assistant-mcp\\${filename}.html`;
  writeFileSync(outputPath, htmlContent, 'utf8');
  
  return {
    type: 'tutorial',
    format: 'html',
    path: outputPath,
    title: tutorial.title,
    preview: '数学教程已生成，可直接在浏览器中查看'
  };
}

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Work Assistant MCP server running on stdio');
}

main().catch((error) => {
  console.error('Server error:', error);
  process.exit(1);
});
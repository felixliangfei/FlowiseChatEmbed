import { EventStreamContentType, fetchEventSource } from '@microsoft/fetch-event-source';
import { z } from "zod";
import { zodToJsonSchema } from "zod-to-json-schema";
import $Q from 'jquery';
import Cookies from 'js-cookie';

export type EEAgentConfig = {
    baseUrl?: string,
    sseEndpoint?: string,
    pushEndpoint?: string,
    maxRetries?: number
}

type FunctionScreenFieldValueType = string | number | Record<string, any>;

const FunctionScreenFieldValueSchema: z.ZodType<FunctionScreenFieldValueType> = z.lazy(() =>
    z.union([
        z.string(),
        z.number(),
        z.record(z.string(), FunctionScreenFieldValueSchema)
    ])
);

const FunctionScreenFieldsSchema = z.array(
    z.object({
        name: z.string()
            .describe("field id or name")
            .min(1),
        description: z.string()
            .describe("field label or description")
            .min(1),
        value: FunctionScreenFieldValueSchema
            .describe("field value")
            .optional(),
    })
);

const ContextSchema = z.object({
    country: z.string().describe("country of the user").optional(),
    unitCode: z.string().describe("unit code of the user").optional(),
    userId: z.string().describe("user id").optional(),
    sessionId: z.string().describe("session id").optional(),
    module: z.string().describe("module of the current function").optional(),
    functionId: z.string().describe("current function id").optional(),
    functionName: z.string().describe("current function name").optional(),
    functionDescription: z.string().describe("current function description").optional(),
    functionShortName: z.string().describe("current function short name").optional(),
    originalFunctionId: z.string().describe("current original function id").optional(),
    originalFunctionName: z.string().describe("current original function name").optional(),
    originalFunctionDescription: z.string().describe("current original function description").optional(),
    originalFunctionShortName: z.string().describe("current original function short name").optional(),
    isFunctionListShow: z.boolean().describe("whether the function list or dashboard is show").optional(),
    isCataScreenShow: z.boolean().describe("whether the catalog screen is show").optional(),
    isFuncScreenShow: z.boolean().describe("whether the function screen is show").optional(),
    functionScreenFieldsSchema: FunctionScreenFieldsSchema.optional(),
    // functionList: z.array(z.object({
    //     functionId: z.string().describe("function id in the function list"),
    //     functionDescription: z.string().describe("function description in the function list"),
    //     functionGroupId: z.string().describe("function group id in the function list"),
    //     functionGroupDescription: z.string().describe("function group description in the function list"),
    // })).describe("function list of the user").optional(),
}).catchall(z.any());
export type ContextType = z.infer<typeof ContextSchema>;
const CommandResultSchema = z.object({
    userId: z.string().describe("user id"),
    sessionId: z.string().describe("session id").optional(),
    result: z.any().describe("result of the command").optional(),
    code: z.number().describe("code of the command").default(0),
    message: z.string().describe("message of the command").default("Completed"),
    schema: z.any().describe("result schema of the command").optional(),
});
export type CommandResultType = z.infer<typeof CommandResultSchema>;

export class EEFrontAgent {
    private readonly config: EEAgentConfig;
    private retryCount: number = 0;
    private readonly baseUrl: string;
    private readonly sseEndpoint: string;
    private readonly pushEndpoint: string;
    private readonly maxRetries: number;
    private readonly commandHandlers = new Map<string, (args: any) => Promise<CommandResultType>>();

    constructor(config: EEAgentConfig) {
        this.config = config;
        this.baseUrl = config.baseUrl ?? "http://localhost:8080/eeagent";
        this.sseEndpoint = config.sseEndpoint ?? "/sse";
        this.pushEndpoint = config.pushEndpoint ?? "/feedback";
        this.maxRetries = config.maxRetries ?? 3;
        this.connectSSE();
    }

    public addCommand(command: string, handler: (args: any) => Promise<CommandResultType>): void {
        this.commandHandlers.set(command, handler);
    }

    public async runCommand(command: string, args: any): Promise<CommandResultType> {
        const handler = this.commandHandlers.get(command);
        if (handler) {
            return await handler(args);
        }
        return { userId: args.headers.userId, sessionId: args.headers.sessionId, code: 404, message: `'${command}' command not found` } as CommandResultType;
    }

    public clearCommands: () => void = () => {
        this.commandHandlers.clear();
    }

    private async onCommand(command: string, args: any): Promise<void> {
        const handler = this.commandHandlers.get(command);
        if (handler) {
            try {
                console.log(`command '${command}' input: ${JSON.stringify(args)}`);
                const result = await handler(args);
                if (result) {
                    this.pushback(command, { ...result, code: 0, message: "Completed" });
                }
            } catch (e: unknown) {
                let message: string;
                if (e instanceof Error) {
                    message = `Error: ${e.message}\n${e.stack}`;
                } else {
                    message = `Unknown error: ${JSON.stringify(e)}`;
                }
                console.error(message);
                this.pushback(command, {
                    userId: args.headers.userId,
                    sessionId: args.headers.sessionId,
                    code: 500,
                    message
                });
            }
        } else {
            this.pushback(command, {
                userId: args.headers.userId,
                sessionId: args.headers.sessionId,
                code: 404,
                message: `'${command}' command not found`
            });
        }
    };
    private readonly pushback: (command: string, data: CommandResultType) => void = (_command, data) => {
        console.log(`command '${_command}' output: ${JSON.stringify(data)}`);
        fetch(this.baseUrl + this.pushEndpoint, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(data)
        }).then(result => console.log(`command '${_command}' output pushback to MCP server completed`))
            .catch(err => console.error(`command '${_command}' output pushback to MCP server error`, err));
    }

    private connectSSE() {
        fetchEventSource(this.config.baseUrl + this.sseEndpoint, {
            openWhenHidden: true,
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                userId: EEUIUtils.getContextData().userId,
                sessionId: EEUIUtils.getContextData().sessionId,
            }),
            async onopen(response) {
                if (response.ok && response.headers.get('content-type')?.startsWith(EventStreamContentType)) {
                    console.log("sse ee agent server connection ok"); // everything's good
                } else if (response.status === 429) {
                    const errMessage = (await response.text()) ?? 'Too many requests. Please try again later.';
                    throw new Error(errMessage);
                } else if (response.status === 403) {
                    const errMessage = (await response.text()) ?? 'Unauthorized';
                    throw new Error(errMessage);
                } else if (response.status === 401) {
                    const errMessage = (await response.text()) ?? 'Unauthenticated';
                    throw new Error(errMessage);
                } else {
                    throw new Error('sse connection error');
                }
            },
            onmessage: async (e) => {
                if (e.event == "command") {
                    console.log(`event message: ${JSON.stringify(e)}`);
                    const { command, input } = JSON.parse(e.data);
                    await this.onCommand(command, input);
                }
            },
            onerror: (err) => {
                console.log(`event message: ${JSON.stringify(err)}`);
                while (this.retryCount < this.maxRetries) {
                    try {
                        this.connectSSE();
                        this.retryCount = 0;
                        break;
                    } catch (e) {
                        if (e) {
                            this.retryCount++;
                        }
                    }
                }
            },
        });
    }

    public disconnect() {
        this.retryCount = this.maxRetries;
    }
}

export class EEUIUtils {
    private constructor() { }

    public static getFrameSet(frameSetId: string, doc?: Document) {
        const $E = EEHtml ?? parent.EEHtml;
        return $E?.getFrameSet(frameSetId, doc);
    }

    public static getFrameWindow(name: string, baseWin?: Window) {
        const $E = EEHtml ?? parent.EEHtml;
        return $E?.getFrameWindow(name, baseWin);
    }

    public static getFrameDocument(name: string, baseWin?: Window) {
        const $E = EEHtml ?? parent.EEHtml;
        return $E?.getFrameDocument(name, baseWin);
    }

    public static getIFrame(iframeId: string, doc?: Document) {
        const $E = EEHtml ?? parent.EEHtml;
        return $E?.getIFrame(iframeId, doc);
    }

    public static getTopWindow() {
        const $E = EEHtml ?? parent.EEHtml;
        return $E?.getTopWindow();
    }
    public static getTopDocument() {
        return EEUIUtils.getTopWindow().document;
    }

    public static getOpenedWindow(name: string) {
        const workWin = EEUIUtils.getFrameWindow("work");
        const childWindows = workWin["arChildWin"] as Array<Window>;
        return childWindows.find(w => w.name === name);
    }

    public static getCSRF() {
        return CSRF;
    }

    public static async sleep(ms: number) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    public static getContextData(): ContextType {
        const workWin = EEUIUtils.getFrameWindow("work") ?? window;
        const funcWin: any = window;
        const isCataScreen = $Q("form", EEUIUtils.getFrameDocument("work")).attr("name") == "cataform";
        const isFuncScreen = $Q("form", EEUIUtils.getFrameDocument("work")).attr("name") == "MAINFORM";
        const generateFuncScreenFieldsSchema = () => {
            return isFuncScreen ? $Q("input,select,textarea", EEUIUtils.getFrameDocument("work")).filter((i, e) => $Q(e).attr("type") != "hidden")
                .map((i, e) => ({
                    name: $Q(e).attr("id") ?? $Q(e).attr("name") ?? "",
                    description: $Q(e).attr("title") ?? "",
                })).filter((i, e) => e.name != "").toArray() : [];
        };
        // const funcList = funcWin["userFunctionList"] as Array<Array<string>>;
        // const funcItems = funcList?.map(i => {
        //     return {
        //         functionId: i[2],
        //         functionDescription: i[1],
        //         functionGroupId: i[5],
        //         functionGroupDescription: i[3]
        //     }
        // });
        return {
            country: workWin["SYS_BANK_COUNTRY"],
            unitCode: workWin["SYS_BUSI_UNIT"],
            userId: funcWin["user"].name ?? workWin["SYS_USER_ID"],
            sessionId: Cookies.get("JSESSIONID"),
            module: workWin["SYS_MODULE_NAME"],
            functionId: workWin["SYS_FUNCTION_ID"],
            functionName: workWin["SYS_FUNCTION_NAME"],
            functionDescription: workWin["SYS_FUNCTION_DESC"],
            functionShortName: workWin["SYS_FUNCTION_SHORT_NAME"],
            originalFunctionId: workWin["SYS_ORG_FUNCTION_ID"],
            originalFunctionName: workWin["SYS_ORG_FUNCTION_NAME"],
            originalFunctionDescription: workWin["SYS_ORG_FUNCTION_DESC"],
            originalFunctionShortName: workWin["SYS_ORG_FUNCTION_SHORT_NAME"],
            isFunctionListShow: workWin["SYS_FUNCTION_ID"] == null || workWin["SYS_FUNCTION_ID"] == "",
            isCataScreenShow: isCataScreen,
            isFuncScreenShow: isFuncScreen,
            functionScreenFieldsSchema: generateFuncScreenFieldsSchema(),
        }
    }

    public static gotoFunction(functionId: string) {
        const operations = window['Operations'] || {};
        operations.Dashboard?.openFunction(functionId); //由于EEV6 UIUX是react组件，外部无法控制react组件中的方法，需要修改Dashboard暴露openFunction方法。
    }
}

export function EEFrontCommandsRegister(agent: EEFrontAgent) {
    agent.addCommand("context.retrieveContext", async args => {
        const ctx = EEUIUtils.getContextData();
        return { userId: args.headers.userId, sessionId: args.headers.sessionId, result: ctx, schema: zodToJsonSchema(ContextSchema, "ContextSchema") } as CommandResultType;
    });
    agent.addCommand("ui.openFunction", async args => {
        const functionId = args.body.functionId;
        EEUIUtils.gotoFunction(functionId);
        return { userId: args.headers.userId, sessionId: args.headers.sessionId, result: {} } as CommandResultType;
    });
    agent.addCommand("ui.openFunctionTemplateList", async args => {
        const loadTmplButton = $Q("#_LoadTmpl", EEUIUtils.getFrameDocument('eeToolbar'));
        if (loadTmplButton.length > 0) {
            $Q("#_LoadTmpl", EEUIUtils.getFrameDocument('eeToolbar')).trigger("click");
        } else {
            $Q("#work").one("load", async () => {
                await EEUIUtils.sleep(1000);
                $Q("#_LoadTmpl", EEUIUtils.getFrameDocument('eeToolbar')).trigger("click");
            });
        }
        return { userId: args.headers.userId, sessionId: args.headers.sessionId, result: {} } as CommandResultType;
    });
    agent.addCommand("ui.searchFunctionTemplateList", async args => {
        // to do
        return { userId: args.headers.userId, sessionId: args.headers.sessionId, result: {} } as CommandResultType;
    });
    agent.addCommand("ui.fillFunctionScreen", async args => {
        // to do
        return { userId: args.headers.userId, sessionId: args.headers.sessionId, result: {} } as CommandResultType;
    });
}

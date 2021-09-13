import * as React from "react";
import { Provider, Flex, Alert, Form, FormInput, FormButton, Grid, Box, Header, mergeThemes, FlexItem, Button } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import Axios from "axios";
import { Provider as RTProvider, CommunicationOptions, List, TListInteraction, TToolbarInteraction } from "@fluentui/react-teams";
import { TeamsTheme } from "@fluentui/react-teams/lib/cjs/themes";
import jwtDecode from "jwt-decode";
import { OnlineMeeting } from "@microsoft/microsoft-graph-types-beta";

enum Errors {
    NotInTeams,
    SSOError,
    NotInTeamsMeeting,
    NoMeetingDetails,
    UnsupportedFrameContext
}

const ErrorMessages: {[key in Errors]: string} = {
    [Errors.NotInTeams]: "This app only works when ran inside of Microsoft Teams.",
    [Errors.SSOError]: "An SSO error occurred.",
    [Errors.NotInTeamsMeeting]: "This app only works in the context of a Microsoft Teams meeting.",
    [Errors.NoMeetingDetails]: "Couldn't determine the meeting details.",
    [Errors.UnsupportedFrameContext]: "This app is not supported in the current FrameContext"
};

/**
 * Implementation of the Meeting Bingo content page
 */
export const MeetingBingoTab = () => {
    const defaultBingoTopics = [
        "Sorry I was on mute",
        "Can you see my screen?",
        "Let's wait for everyone to join",
        "Dog barks",
        "Cat meows",
        "Kids show up",
        "Wrong screen shared 😱",
        "I have a hard stop at",
        "I can't access the chat",
        "Leaves virtual hand raised",
        "I have to jump to another call",
        "Let me share my screen",
        "Can you hear me?",
        "Oh I thought my camera was off",
        "Someone drinks coffee",
        "Someone types loudly",
        "Someone yawns ",
        "Someone is eating",
        "Intense chat is happening during presentation",
        "has x joined?",
        "Sorry I'm late / what is this about?",
        "Someone forgets to stop sharing their screen"
    ];
    const [{ inTeams, theme, themeString, context }] = useTeams();
    const [errors, setErrors] = useState<Errors[]>([]);
    const [meetingId, setMeetingId] = useState<string | undefined>();
    const [accessToken, setAccessToken] = useState<string>();
    const [frameContext, setFrameContext] = useState<microsoftTeams.FrameContexts | null>();
    const [onlineMeeting, setOnlineMeeting] = useState<OnlineMeeting>({});
    const [bingoTopics, setBingoTopics] = useState<string[]>([]);
    const [currentUserId, setcurrentUserId] = useState<string>();
    const [newTopicValue, setNewTopicValue] = useState<string>();
    const [showAddTopicForm, setShowAddTopicForm] = useState<boolean>(false);
    const [bingoGrid, setBingoGrid] = useState<{ selected: boolean; value: string; }[][]>([[]]);

    const addError = (errs: Errors[], err: Errors) => {
        return [...errs, err];
    };

    const removeError = (errs: Errors[], err: Errors) => {
        return errs.filter((i) => {
            return i !== err;
        });
    };

    const getOrCreateBingoGrid = (meetingId: string, force = false) => {
        const bingoGrid = localStorage.getItem(meetingId);
        if (!bingoGrid || force) {
            const newBingoGridItems = [...bingoTopics].sort(() => 0.5 - Math.random()).slice(0, 9);
            const newBingoGrid: { selected: boolean; value: string; }[][] = [];
            while (newBingoGridItems.length > 0) {
                newBingoGrid.push(newBingoGridItems.splice(0, 3).map((v) => ({
                    selected: false,
                    value: v
                })));
            }
            localStorage.setItem(meetingId, JSON.stringify(newBingoGrid));
            setBingoGrid(newBingoGrid);
        } else {
            setBingoGrid(JSON.parse(bingoGrid) as { selected: boolean; value: string; }[][]);
        }
    };

    const selectBingoGridCell = async (row: number, column: number) => {
        const newBingoGrid = [...bingoGrid];
        newBingoGrid[row][column].selected = !newBingoGrid[row][column].selected;
        localStorage.setItem(meetingId as string, JSON.stringify(newBingoGrid));
        setBingoGrid(newBingoGrid);

        if ((newBingoGrid[row][0].selected && newBingoGrid[row][1].selected && newBingoGrid[row][2].selected) || (newBingoGrid[0][column].selected && newBingoGrid[1][column].selected && newBingoGrid[2][column].selected)) {
            await Axios.post(`https://${process.env.PUBLIC_HOSTNAME}/api/chatMessage/${meetingId}`, { body: { content: "BINGO!" } }, { headers: { Authorization: `Bearer ${accessToken}` } });
        }
    };

    useEffect(() => {
        if (inTeams === true) {
            setErrors(errs => removeError(errs, Errors.NotInTeams));
            microsoftTeams.authentication.getAuthToken({
                successCallback: (token: string) => {
                    const decoded: { [key: string]: any; } = jwtDecode(token) as { [key: string]: any; };
                    setAccessToken(token);
                    setcurrentUserId(decoded.oid);
                    setErrors(errs => removeError(errs, Errors.SSOError));
                    microsoftTeams.appInitialization.notifySuccess();
                },
                failureCallback: (message: string) => {
                    setErrors(errs => addError(errs, Errors.SSOError));
                    microsoftTeams.appInitialization.notifyFailure({
                        reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
                        message
                    });
                },
                resources: [`api://${process.env.PUBLIC_HOSTNAME}/${process.env.TAB_APP_ID}`]
            });
        } else {
            setErrors(errs => [...errs, Errors.NotInTeams]);
        }
    }, [inTeams]);

    useEffect(() => {
        if (context) {
            setMeetingId(context.meetingId);
            setFrameContext(context.frameContext);

            if (!context.meetingId) {
                setErrors(errs => addError(errs, Errors.NotInTeamsMeeting));
            } else {
                setErrors(errs => removeError(errs, Errors.NotInTeamsMeeting));
            }

            if (context.frameContext === microsoftTeams.FrameContexts.content || context.frameContext === microsoftTeams.FrameContexts.sidePanel || context.frameContext === microsoftTeams.FrameContexts.meetingStage) {
                setErrors(errs => removeError(errs, Errors.UnsupportedFrameContext));
                if (context.meetingId && (context.frameContext === microsoftTeams.FrameContexts.sidePanel || context.frameContext === microsoftTeams.FrameContexts.meetingStage)) {
                    getOrCreateBingoGrid(context.meetingId);
                }
            } else {
                setErrors(errs => addError(errs, Errors.UnsupportedFrameContext));
            }
        }
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [context]);

    useEffect(() => {
        (async () => {
            if (meetingId && accessToken) {
                const response = await Axios.get<OnlineMeeting>(`https://${process.env.PUBLIC_HOSTNAME}/api/meetingDetails/${meetingId}`, { headers: { Authorization: `Bearer ${accessToken}` } });
                setOnlineMeeting(response.data);

                if (!response.data) {
                    setErrors(errs => addError(errs, Errors.NoMeetingDetails));
                } else {
                    setErrors(errs => removeError(errs, Errors.NoMeetingDetails));
                }

                const storedTopics = await Axios.get<string[]>(`https://${process.env.PUBLIC_HOSTNAME}/api/bingoTopics/${meetingId}`, { headers: { Authorization: `Bearer ${accessToken}` } });
                setBingoTopics(storedTopics.data);
            }
        })();
    }, [meetingId, accessToken]);

    const updateTopics = async (newTopics: string[]) => {
        await Axios.post<string[]>(`https://${process.env.PUBLIC_HOSTNAME}/api/bingoTopics/${meetingId}`, newTopics, { headers: { Authorization: `Bearer ${accessToken}` } });
        setBingoTopics(newTopics);
    };

    let mainContent: JSX.Element | JSX.Element[] | null = null;
    if (errors.length > 0) {
        mainContent = errors.map((err) => <Alert
            key={err}
            content={ErrorMessages[err]}
            variables={{
                urgent: true
            }}
        />);
    } else if (frameContext === microsoftTeams.FrameContexts.content) {
        let initActions;
        let topicAction;
        let addTopicAction = {};
        let addTopicForm: JSX.Element | null = null;
        let gridSpan = {
            gridColumn: "span 4"
        };
        let selectable = false;
        if (onlineMeeting && currentUserId && onlineMeeting?.participants?.organizer?.identity?.user?.id === currentUserId) {
            initActions = {
                primary: {
                    label: "Initialize Meeting Bingo topics",
                    target: "initTopics"
                }
            };
            topicAction = {
                deleteTopic: {
                    title: "Remove topic",
                    icon: "TrashCan",
                    multi: true
                }
            };
            addTopicAction = {
                g1: {
                    addTopic: {
                        title: "Add topic",
                        icon: "Add"
                    }
                }
            };
            selectable = true;
            if (showAddTopicForm) {
                gridSpan = {
                    gridColumn: "span 3"
                };
                addTopicForm = <Provider theme={theme}>
                    <Box styles={{
                        gridColumn: "span 1"
                    }}>
                        <Flex fill={true} column styles={{
                            paddingLeft: "1.6rem",
                            paddingRight: "1.6rem"
                        }}>
                            <Header content="Add topic" />
                            <Form
                                styles={{
                                    justifyContent: "initial"
                                }}
                                onSubmit={() => {
                                    if (newTopicValue) {
                                        updateTopics([...bingoTopics, newTopicValue]);
                                    }
                                    setNewTopicValue("");
                                }}
                            >
                                <FormInput
                                    label="Topic"
                                    name="topic"
                                    id="topic"
                                    required
                                    value={newTopicValue}
                                    onChange={(e, i) => {
                                        setNewTopicValue(i?.value);
                                    }}
                                    showSuccessIndicator={false}
                                />
                                <FormButton content="Submit" primary />
                            </Form>
                            <FlexItem push>
                                <Button content="Close" secondary onClick={() => { setShowAddTopicForm(false); }} style={{
                                    marginLeft: "auto",
                                    marginRight: "auto",
                                    marginTop: "2rem",
                                    width: "12rem"
                                }} />
                            </FlexItem>
                        </Flex>
                    </Box>
                </Provider>;
            }
        }

        const rows = bingoTopics.map(value => ({
            topic: value,
            actions: topicAction
        })).reduce((pv, cv, i, arr) => ({ ...pv, [i.toString()]: cv }), {});

        mainContent = <Grid columns="repeat(4, 1fr)" styles={{
            gap: "20px"
        }}>
            <Box styles={gridSpan} >
                <Flex fill={true} column>
                    <List
                        emptyState={{
                            fields: {
                                title: "Create your first Meeting Bingo topic",
                                desc: "Get started by adding your own topics or initialize with a default set of topics",
                                actions: initActions
                            },
                            option: CommunicationOptions.Empty
                        }}
                        emptySelectionActionGroups={addTopicAction}
                        columns={{
                            topic: {
                                title: "Topic"
                            }
                        }}
                        rows={rows}
                        onInteraction={async (interaction: TListInteraction) => {
                            if (interaction.target === "initTopics") {
                                await updateTopics([...defaultBingoTopics]);
                            } else if (interaction.target === "toolbar") {
                                const toolbarInteraction = interaction as TToolbarInteraction;
                                if (toolbarInteraction.action === "deleteTopic") {
                                    if (Array.isArray(toolbarInteraction.subject)) {
                                        const indexes = toolbarInteraction.subject.map(v => parseInt(v));
                                        indexes.sort((a, b) => b - a);
                                        indexes.forEach(v => bingoTopics.splice(v, 1));
                                        await updateTopics([...bingoTopics]);
                                    }
                                } else if (toolbarInteraction.action === "addTopic") {
                                    setShowAddTopicForm(true);
                                }
                            }
                        }}
                        selectable={selectable}
                        title="Meeting Bingo" />
                </Flex>
            </Box>
            {addTopicForm}
        </Grid>;
    } else if (frameContext === microsoftTeams.FrameContexts.sidePanel || frameContext === microsoftTeams.FrameContexts.meetingStage) {
        let content: JSX.Element | null = null;
        if (bingoTopics.length <= 8) {
            content = <Alert
                content="Not enough Meeting Bingo topics are configured to play the game."
                variables={{
                    urgent: true
                }}
            />;
        } else {
            const rowContent = bingoGrid.map((v, r) => <tr key={`${r}`}>
                {v.map((i, c) => <td key={`${r}-${c}`} data-item={JSON.stringify({ row: r, column: c })} style={{
                    padding: "10px",
                    textAlign: "center",
                    verticalAlign: "middle",
                    backgroundColor: (i.selected) ? theme.siteVariables.colorScheme.brand.background1 : "inherit",
                    cursor: "pointer",
                    border: `1px solid ${theme.siteVariables.colorScheme.brand.background1}`
                }} onClick={async (e) => {
                    const itemString = (e.target as HTMLTableCellElement).getAttribute("data-item");
                    if (itemString) {
                        const item = JSON.parse(itemString);
                        await selectBingoGridCell(item.row, item.column);
                    }
                }}>
                    {i.value}
                </td>)}
            </tr>);

            content = <React.Fragment>
                <table style={{
                    marginLeft: "auto",
                    marginRight: "auto",
                    marginTop: "2rem",
                    marginBottom: "2rem",
                    borderCollapse: "collapse",
                    border: `1px solid ${theme.siteVariables.colorScheme.brand.background1}`
                }}>
                    {rowContent}
                </table>
                <Button content="Refresh bingo card" secondary onClick={() => { getOrCreateBingoGrid(meetingId as string, true); }} style={{
                    marginLeft: "auto",
                    marginRight: "auto",
                    marginTop: "2rem",
                    marginBottom: "2rem",
                    width: "12rem"
                }} />
            </React.Fragment>;
        }
        mainContent = <Flex fill={true} column>
            {content}
        </Flex>;
    }

    /**
     * The render() method to create the UI of the tab
     */
    return (
        <Provider theme={theme}>
            <RTProvider themeName={TeamsTheme[themeString.charAt(0).toUpperCase() + themeString.slice(1)]} lang="en-US">
                {mainContent}
            </RTProvider>
        </Provider>
    );
};

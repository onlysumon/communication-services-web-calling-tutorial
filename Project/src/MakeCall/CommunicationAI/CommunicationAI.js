import React, { useState, useEffect } from "react";
import { Features, ResultTyp, CallKind } from '@azure/communication-calling';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import { utils, acsOpenAiPromptsApi } from "../../Utils/Utils";
import { PrimaryButton
} from 'office-ui-fabric-react'
import HtmlParser from "react-html-parser";


const CommunicationAI = ({ call }) => {
    const [captionsStarted, setCaptionsStarted] = useState(false);
    const [showSpinner, setShowSpinner] = useState(false);
    const [captionHistory, setCaptionHistory] = useState([]);
    const [lastSummary, setLastSummary] = useState("");
    const [captionsSummaryIndex, setCaptionsSummaryIndex] = useState(0);
    const [lastFeedBack, setLastFeedBack] = useState("");
    const [captionsFeedbackIndex, setCaptionsFeedbackIndex] = useState(0);
    const [promptResponse, setPromptResponse] = useState("")
    const [dropDownLabel, setDropDownLabel] = useState("")
    const [showTaskButton, setShowTaskButton] = useState(false);
    const [isSpeaking, setIsSpeaking] = useState(true);
    const [feedBackMessage, setFeedBackMessage] = useState("");
    const [debounceCounterRunning, setDebounceCounterRunning] = useState(false);


    const options = [
        { key: 'getSummary', text: 'Get Summary' },
        { key: 'getPersonalFeedBack', text: 'Get Personal Feedback' },
    ]
    const displayName = window.displayName;
    let captions;
    let debounceTimeout;
    useEffect(() => {
        captions = call.kind === CallKind.TeamsCall || call.info?.context === 'teamsMeetingJoin' ? call.feature(Features.TeamsCaptions) : call.feature(Features.Captions);
        startCaptions(captions);

        return () => {
            // cleanup
            captions.off('captionsReceived', captionHandler);
        };
    }, []);

    useEffect(() => {
        if (dropDownLabel == 'getPersonalFeedBack' && !isSpeaking) {
            debounceCounterRunning && clearTimeout(debounceTimeout);
            console.log(`Starting debounce timer`);
            setDebounceCounterRunning(true);
            debounceTimeout = setTimeout(() => {
                setDebounceCounterRunning(false);
                getPersonalFeedback()}, 5000);
             return () => {
                clearTimeout(debounceTimeout);
              };
        } else {
            setFeedBackMessage('FeedBack will be retrieved after you finish talking')
        }
    }, [isSpeaking]);

    const startCaptions = async () => {
        try {
            if (!captions.isCaptionsActive || !captionsStarted) {
                await captions.startCaptions({ spokenLanguage: 'en-us' });
                setCaptionsStarted(!captionsStarted);
            }
            captions.on('captionsReceived', captionHandler);
        } catch (e) {
            console.error('startCaptions failed', e);
        }
    };

    const captionHandler = (captionData) => {
        let mri = '';
        if (captionData.speaker.identifier.kind === 'communicationUser') {
            mri = captionData.speaker.identifier.communicationUserId;
        } else if (captionData.speaker.identifier.kind === 'microsoftTeamsUser') {
            mri = captionData.speaker.identifier.microsoftTeamsUserId;
        } else if (captionData.speaker.identifier.kind === 'phoneNumber') {
            mri = captionData.speaker.identifier.phoneNumber;
        }
        mri == window.mri && setIsSpeaking(true)
        const captionText = `${captionData.speaker.displayName}: ${captionData.text}`;

        if (captionData.resultType === ResultType.Final) {
            setCaptionHistory(oldCaptions => [...oldCaptions, captionText]);
            mri == window.mri && setIsSpeaking(false)
        }
    };


    const getSummary = async () => {
        const currentCaptionsData = captionHistory.slice(captionsSummaryIndex);
        let response = await utils.sendCaptionsDataToAcsOpenAI(acsOpenAiPromptsApi.summary, displayName, lastSummary, currentCaptionsData);
        const content = response.choices[0].message.content;
        setLastSummary(content);
        setCaptionsSummaryIndex(captionHistory.length);
        setPromptResponse(content.split("\n").join("<br />"));
    }

    const getPersonalFeedback = async () => {
        setShowSpinner(true)
        setFeedBackMessage('Retrieving Personal FeedBack')
        const currentCaptionsData = captionHistory.slice(captionsFeedbackIndex);
        let response = await utils.sendCaptionsDataToAcsOpenAI(acsOpenAiPromptsApi.feedback, displayName, lastFeedBack, currentCaptionsData)
        const content = response.choices[0].message.content

        setLastFeedBack(content);
        setCaptionsFeedbackIndex(captionHistory.length);
        setPromptResponse(content.split("\n").join("<br />"));
        setShowSpinner(false)
        setShowTaskButton(true);
        setIsSpeaking(false)
        setShowActionItem(true);
    }

    const onChangeHandler = (e, item) => {
        let communicationAiOption = item.key;
        setDropDownLabel(communicationAiOption);
        setShowSpinner(true);
        switch (communicationAiOption) {
            case "getSummary":
                getSummary().finally(() => setShowSpinner(false));
                break;
            case "getPersonalFeedBack":
                break;
        }

    }

    const createTasks = async () => {
        setShowSpinner(true)
        setShowTaskButton(false)
        let response = await utils.sendCaptionsDataToAcsOpenAI(acsOpenAiPromptsApi.feedbackTasks, displayName,"", lastFeedBack);
        const content = response.choices[0].message.content;
        setActionItems(content.split("\n").join("<br />"));
        setShowSpinner(false)
        setShowTaskButton(false)
    }

    return (
        <>
            <div id="" className="">
                <Dropdown
                    placeholder="Select an option"
                    label=""
                    options={options}
                    styles={{ dropdown: { width: 300 }, }}
                    onChange={onChangeHandler}
                />
            </div>

            <div id="communicationResponse">
                {
                    showSpinner &&
                    <div>
                        <div className="loader inline-block"> </div>
                        <div className="ml-2 inline-block">
                            {
                                (dropDownLabel == "getPersonalFeedBack") ?
                                    feedBackMessage :
                                    "Waiting for the AI response..."
                            }
                            </div>
                    </div>
                }
                {showSpinner ? '' : HtmlParser(promptResponse)}
            </div>
        </>
    );
};

export default CommunicationAI;